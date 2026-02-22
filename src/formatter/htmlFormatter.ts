import * as vscode from 'vscode';
import * as prettier from 'prettier';
import { formatSingleAspBlock, getAspSettings, FormatBlockResult } from './aspFormatter';

export interface PrettierSettings {
    printWidth: number;
    tabWidth: number;
    useTabs: boolean;
    semi: boolean;
    singleQuote: boolean;
    bracketSameLine: boolean;
    arrowParens: string;
    trailingComma: string;
    endOfLine: string;
    htmlWhitespaceSensitivity: string;
}

export function getPrettierSettings(): PrettierSettings {
    const config = vscode.workspace.getConfiguration('aspLanguageSupport.prettier');
    return {
        printWidth:                config.get<number>('printWidth', 80),
        tabWidth:                  config.get<number>('tabWidth', 2),
        useTabs:                   config.get<boolean>('useTabs', false),
        semi:                      config.get<boolean>('semi', true),
        singleQuote:               config.get<boolean>('singleQuote', false),
        bracketSameLine:           config.get<boolean>('bracketSameLine', true),
        arrowParens:               config.get<string>('arrowParens', 'always'),
        trailingComma:             config.get<string>('trailingComma', 'es5'),
        endOfLine:                 config.get<string>('endOfLine', 'lf'),
        htmlWhitespaceSensitivity: config.get<string>('htmlWhitespaceSensitivity', 'css'),
    };
}

// Module-level counter so IDs stay unique across calls in the same millisecond.
let _placeholderCounter = 0;

// ─── Types ─────────────────────────────────────────────────────────────────

// Where in the HTML structure an ASP block sits:
//   normal  – standalone on its own line(s)    → HTML comment placeholder
//   inline  – inside a quoted attribute value  → bare token placeholder
//   midtag  – between attributes, not quoted   → data- attribute placeholder
type AspBlockKind = 'normal' | 'inline' | 'midtag';

interface AspBlock {
    code: string;
    indent: string;
    id: string;
    lineNumber: number;
    kind: AspBlockKind;
}

// ─── Safety: detect unclosed ASP tags ─────────────────────────────────────

// An unclosed <% causes the regex to scan to end-of-file, potentially
// consuming and destroying everything after it.
function hasUnclosedAspTags(code: string): boolean {
    let depth = 0;
    let i = 0;
    while (i < code.length) {
        if (code[i] === '<' && code[i + 1] === '%') {
            depth++;
            i += 2;
        } else if (code[i] === '%' && code[i + 1] === '>') {
            if (depth === 0) return true; // stray %>
            depth--;
            i += 2;
        } else {
            i++;
        }
    }
    return depth !== 0;
}

// ─── Classify offset position in the HTML ─────────────────────────────────

// Walks backwards from `offset` to determine where the ASP block sits.
function classifyOffset(code: string, offset: number): AspBlockKind {
    let inQuote: string | null = null;

    for (let i = offset - 1; i >= 0; i--) {
        const ch = code[i];

        if (inQuote) {
            if (ch === inQuote) inQuote = null;
            continue;
        }

        if (ch === '"' || ch === "'") {
            inQuote = ch; // hit closing quote while scanning backwards
            continue;
        }

        if (ch === '>') return 'normal';

        if (ch === '<') {
            const after = code.substring(i + 1, i + 3);
            // Closing tags (</...) can never have attributes — treat as normal.
            if (after.startsWith('/')) return 'normal';
            // Real opening/void tag or declaration.
            if (/^[a-zA-Z!?]/.test(after)) return 'midtag';
            return 'normal';
        }
    }

    return 'normal';
}

// ─── Main formatter ────────────────────────────────────────────────────────

export async function formatCompleteAspFile(code: string): Promise<string> {
    if (hasUnclosedAspTags(code)) {
        console.warn('ASP formatter: unclosed <% or stray %> found, skipping format.');
        return code;
    }

    const aspSettings      = getAspSettings();
    const prettierSettings = getPrettierSettings();

    // ── Step 1: Mask all ASP blocks ──────────────────────────────────────

    const aspBlocks: AspBlock[] = [];

    const maskedCode = code.replace(/([ \t]*)(<%[\s\S]*?%>)/g, (match, indent, aspBlock, offset) => {
        const kind       = classifyOffset(code, offset);
        const lineNumber = code.substring(0, offset).split('\n').length - 1;
        const id         = `ASPPH${_placeholderCounter++}_${Date.now().toString(36)}_${Math.random().toString(36).substring(2, 9)}`;

        aspBlocks.push({ code: aspBlock, indent, id, lineNumber, kind });

        switch (kind) {
            case 'inline': return `ASPINLINE_${id}_END`;
            case 'midtag': return `data-asp-${id}="1"`;
            default:       return `${indent}<!--${id}-->`;
        }
    });

    // ── Step 2: Run Prettier ─────────────────────────────────────────────

    let prettifiedCode: string;
    try {
        prettifiedCode = await prettier.format(maskedCode, {
            parser:                    'html',
            printWidth:                prettierSettings.printWidth,
            tabWidth:                  prettierSettings.tabWidth,
            useTabs:                   prettierSettings.useTabs,
            semi:                      prettierSettings.semi,
            singleQuote:               prettierSettings.singleQuote,
            bracketSameLine:           prettierSettings.bracketSameLine,
            arrowParens:               prettierSettings.arrowParens               as any,
            trailingComma:             prettierSettings.trailingComma             as any,
            endOfLine:                 prettierSettings.endOfLine                 as any,
            htmlWhitespaceSensitivity: prettierSettings.htmlWhitespaceSensitivity as any,
        });
    } catch (error) {
        console.error('ASP formatter: Prettier failed, returning original.', error);
        return code;
    }

    // ── Step 3: Verify all placeholders survived Prettier ────────────────

    for (const block of aspBlocks) {
        const needle =
            block.kind === 'inline' ? `ASPINLINE_${block.id}_END` :
            block.kind === 'midtag' ? `data-asp-${block.id}`       :
            block.id;

        if (!prettifiedCode.includes(needle)) {
            console.error(`ASP formatter: placeholder lost at line ${block.lineNumber}, returning original.`);
            return code;
        }
    }

    // ── Step 4: Format each ASP block's VBScript content ─────────────────
    // Process sequentially so each normal block threads its ending indent
    // level into the next, enabling cross-block indent continuity.

    const formattedBlocks: string[] = [];
    let currentIndentLevel = 0;

    for (const block of aspBlocks) {
        if (block.kind !== 'normal') {
            formattedBlocks.push(block.code);
            continue;
        }
        const result = formatSingleAspBlock(block.code, aspSettings, '', currentIndentLevel);
        formattedBlocks.push(result.formatted);
        currentIndentLevel = result.endLevel;
    }

    // ── Step 5: Restore blocks into the Prettier output ──────────────────

    let restoredCode = prettifiedCode;

    for (let i = 0; i < aspBlocks.length; i++) {
        const block          = aspBlocks[i];
        const formattedBlock = formattedBlocks[i];
        const escapedId      = block.id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

        switch (block.kind) {

            case 'inline':
                restoredCode = restoredCode.replace(
                    `ASPINLINE_${block.id}_END`,
                    formattedBlock
                );
                break;

            case 'midtag':
                // Prettier may have normalised quotes/spacing around the attribute.
                restoredCode = restoredCode.replace(
                    new RegExp(`\\s*data-asp-${escapedId}\\s*=\\s*["']1["']`),
                    ` ${formattedBlock}`
                );
                break;

            default: {
                // Pick up whatever indentation Prettier gave the comment line,
                // then apply it to every non-empty line of the formatted block.
                const match = restoredCode.match(new RegExp(`([ \\t]*)<!--${escapedId}-->`));

                if (match) {
                    const htmlIndent    = match[1];
                    const indentedBlock = formattedBlock
                        .split('\n')
                        .map(line => (line.trim() ? htmlIndent + line : line))
                        .join('\n');

                    restoredCode = restoredCode.replace(
                        new RegExp(`[ \\t]*<!--${escapedId}-->`),
                        indentedBlock
                    );
                } else {
                    // Fallback — sanity check above should have caught real deletions.
                    restoredCode = restoredCode.replace(`<!--${block.id}-->`, formattedBlock);
                }
                break;
            }
        }
    }

    return restoredCode;
}