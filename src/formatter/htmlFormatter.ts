import * as vscode from 'vscode';
import * as prettier from 'prettier';
import { formatSingleAspBlock, getAspSettings, FormatBlockResult } from './aspFormatter';
import { isInsideAspBlock } from '../utils/documentHelper';

// ─── Prettier settings ─────────────────────────────────────────────────────

/**
 * Prettier formatting options surfaced under the
 * `aspLanguageSupport.prettier.*` configuration namespace.
 *
 * HTML, CSS, and JavaScript formatting is delegated entirely to Prettier
 * (https://prettier.io). These settings map 1-to-1 to Prettier's own options.
 */
export interface PrettierSettings {
    printWidth:                number;
    tabWidth:                  number;
    useTabs:                   boolean;
    semi:                      boolean;
    singleQuote:               boolean;
    bracketSameLine:           boolean;
    arrowParens:               string;
    trailingComma:             string;
    endOfLine:                 string;
    htmlWhitespaceSensitivity: string;
}

export function getPrettierSettings(): PrettierSettings {
    const c = vscode.workspace.getConfiguration('aspLanguageSupport.prettier');
    return {
        printWidth:                c.get<number>('printWidth',                80),
        tabWidth:                  c.get<number>('tabWidth',                  2),
        useTabs:                   c.get<boolean>('useTabs',                  false),
        semi:                      c.get<boolean>('semi',                     true),
        singleQuote:               c.get<boolean>('singleQuote',              false),
        bracketSameLine:           c.get<boolean>('bracketSameLine',          true),
        arrowParens:               c.get<string>('arrowParens',               'always'),
        trailingComma:             c.get<string>('trailingComma',             'es5'),
        endOfLine:                 c.get<string>('endOfLine',                 'lf'),
        htmlWhitespaceSensitivity: c.get<string>('htmlWhitespaceSensitivity', 'css'),
    };
}

// ─── ASP block types ───────────────────────────────────────────────────────

// Where in the HTML structure an ASP block sits:
//   normal  – standalone on its own line(s)   → HTML comment placeholder
//   inline  – inside a quoted attribute value → bare token placeholder
//   midtag  – between attributes, not quoted  → data- attribute placeholder
type AspBlockKind = 'normal' | 'inline' | 'midtag';

interface AspBlock {
    code:       string;
    id:         string;
    lineNumber: number;
    kind:       AspBlockKind;
}

// Module-level counter keeps IDs unique across calls in the same millisecond.
let _placeholderCounter = 0;

// ─── Safety check ─────────────────────────────────────────────────────────

/**
 * Returns true if the source has unmatched <% or %> tags.
 * An unclosed <% would cause the masking regex to consume everything after it.
 *
 * Skips:
 *  - HTML comments  <!-- ... -->  entirely (ASP tags inside them are not real)
 *  - VBScript comment lines (first non-whitespace char is ') inside ASP blocks
 *  - %> inside string literals inside ASP blocks
 */
function hasUnclosedAspTags(code: string): boolean {
    let depth   = 0;
    let i       = 0;
    let inHtmlComment = false;

    while (i < code.length) {
        // ── HTML comment open  <!-- ──────────────────────────────────────────
        if (!inHtmlComment && depth === 0 &&
            code[i] === '<' && code.slice(i, i + 4) === '<!--') {
            inHtmlComment = true;
            i += 4;
            continue;
        }
        // ── HTML comment close  --> ──────────────────────────────────────────
        if (inHtmlComment) {
            if (code.slice(i, i + 3) === '-->') { inHtmlComment = false; i += 3; }
            else { i++; }
            continue;
        }

        // ── ASP open  <% ─────────────────────────────────────────────────────
        if (code[i] === '<' && code[i + 1] === '%') {
            depth++;
            i += 2;
            continue;
        }

        // ── Inside ASP block: scan line-by-line ──────────────────────────────
        if (depth > 0) {
            const lineEnd  = code.indexOf('\n', i);
            const lineText = lineEnd === -1 ? code.slice(i) : code.slice(i, lineEnd + 1);
            const end      = lineEnd === -1 ? code.length   : lineEnd + 1;

            // VBScript comment line — no %> on this line counts
            if (lineText.trimStart().startsWith("'")) {
                i = end;
                continue;
            }

            // Scan line for %> outside string literals
            let j = i, inStr = false, found = false;
            while (j < end) {
                if (code[j] === '"') {
                    if (inStr && j + 1 < end && code[j + 1] === '"') { j += 2; continue; }
                    inStr = !inStr; j++; continue;
                }
                if (!inStr && code[j] === '%' && j + 1 < code.length && code[j + 1] === '>') {
                    depth--;
                    if (depth < 0) { return true; }
                    j += 2; found = true; i = j; break;
                }
                j++;
            }
            if (!found) { i = end; }
            continue;
        }

        i++;
    }

    return depth !== 0;
}

// ─── ASP block classifier ─────────────────────────────────────────────────

/**
 * Walks backwards from `offset` in the original source to determine whether
 * the ASP block at that position is inside a quoted attribute value (inline),
 * between unquoted attributes (midtag), or free-standing (normal).
 */
function classifyOffset(code: string, offset: number): AspBlockKind {
    let inQuote: string | null = null;
    let i = offset - 1;

    while (i >= 0) {
        const ch = code[i];

        if (inQuote) {
            if (ch === inQuote) inQuote = null;
            i--; continue;
        }

        if (ch === '"' || ch === "'") {
            inQuote = ch; // hit closing quote while scanning backwards
            i--; continue;
        }

        if (ch === '>') {
            // If this > is part of a %> closing tag, it is NOT a real HTML tag
            // boundary — skip backwards past the entire <% ... %> block and keep
            // scanning. Without this guard, a pattern like:
            //   <option <% If x Then %>selected<% End If %>>
            // would make the second <% (End If) see the %> from the first block
            // and incorrectly return 'normal', causing an HTML comment placeholder
            // to be injected mid-tag and breaking Prettier.
            if (i > 0 && code[i - 1] === '%') {
                i -= 2; // step past %>
                while (i >= 0) {
                    if (code[i] === '<' && i + 1 < code.length && code[i + 1] === '%') {
                        i--; // step past <%, continue outer loop
                        break;
                    }
                    i--;
                }
                continue;
            }
            return 'normal'; // real HTML tag close — we are between tags
        }

        if (ch === '<') {
            const after = code.substring(i + 1, i + 3);
            if (after.startsWith('/'))        return 'normal';  // closing tag
            if (/^[a-zA-Z!?]/.test(after))   return 'midtag';  // opening/void tag
            return 'normal';
        }

        i--;
    }

    return 'normal';
}

// ─── Main entry point ──────────────────────────────────────────────────────

export async function formatCompleteAspFile(code: string): Promise<string> {
    if (hasUnclosedAspTags(code)) {
        vscode.window.showWarningMessage(
            'Formatting skipped — unclosed <% or stray %> detected. Fix the ASP tag mismatch first.'
        );
        return code;
    }

    const aspSettings      = getAspSettings();
    const prettierSettings = getPrettierSettings();

    // ── Step 1: Mask all ASP blocks ──────────────────────────────────────────
    // Each ASP block is replaced with a placeholder that Prettier will treat as
    // valid HTML, preserving its position in the output.

    const aspBlocks: AspBlock[] = [];

    // Manual scan so we can skip <% ... %> blocks that sit inside an HTML
    // comment <!-- ... -->.  A regex replace has no context awareness.
    let   maskedCode     = '';
    let   pos            = 0;
    let   inHtmlComment  = false;

    while (pos < code.length) {
        // ── HTML comment open  <!-- ────────────────────────────────────────
        if (!inHtmlComment && code.slice(pos, pos + 4) === '<!--') {
            // Before entering the comment check it's not an ASP placeholder we
            // already emitted (those start with <!--ASPPH) — shouldn't happen
            // here but guard anyway.
            inHtmlComment = true;
            maskedCode   += code[pos];
            pos++;
            continue;
        }
        // ── HTML comment close  --> ────────────────────────────────────────
        if (inHtmlComment) {
            if (code.slice(pos, pos + 3) === '-->') { inHtmlComment = false; }
            maskedCode += code[pos];
            pos++;
            continue;
        }

        // ── ASP block  <% ... %> ──────────────────────────────────────────
        if (code[pos] === '<' && code[pos + 1] === '%') {
            // Collect any leading horizontal whitespace on this line for indent tracking
            const lineStart   = maskedCode.lastIndexOf('\n') + 1;
            const leadingWS   = maskedCode.slice(lineStart).match(/^[ \t]*/)?.[0] ?? '';

            // Find the matching %> (using the same comment-aware logic so a '
            // comment line inside the block can't close it prematurely)
            let end = pos + 2;
            while (end < code.length) {
                if (code[end] === '%' && code[end + 1] === '>') {
                    // Check whether this %> is on a VBScript comment line
                    const lineBegin = code.lastIndexOf('\n', end - 1) + 1;
                    const lineUpToClose = code.slice(lineBegin, end).trimStart();
                    if (!lineUpToClose.startsWith("'")) {
                        end += 2; // include the %>
                        break;
                    }
                }
                end++;
            }

            const aspBlock   = code.slice(pos, end);
            const kind       = classifyOffset(code, pos);
            const lineNumber = code.slice(0, pos).split('\n').length - 1;
            const id         = `ASPPH${_placeholderCounter++}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;

            aspBlocks.push({ code: aspBlock, id, lineNumber, kind });

            // Replace leading whitespace + block with the placeholder
            // (strip the leadingWS we already emitted so the placeholder
            //  takes its place cleanly)
            if (leadingWS && maskedCode.endsWith(leadingWS)) {
                maskedCode = maskedCode.slice(0, maskedCode.length - leadingWS.length);
            }

            switch (kind) {
                case 'inline': maskedCode += `ASPINLINE_${id}_END`; break;
                case 'midtag': maskedCode += `data-asp-${id}="1"`; break;
                default:       maskedCode += `<!--${id}-->`; break;
            }
            pos = end;
            continue;
        }

        maskedCode += code[pos];
        pos++;
    }

    // ── Step 2: Run Prettier on the masked HTML ──────────────────────────────

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
            // Prevent Prettier from reformatting JS inside inline event handlers
            // (onchange="a(); b()") and CSS inside style="".  Without this,
            // Prettier expands multi-statement handlers onto separate indented
            // lines when the tag wraps, which is not what we want.
            embeddedLanguageFormatting: 'off' as any,
        });
    } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        const lineMatch = msg.match(/\((\d+):(\d+)\)/);
        const location  = lineMatch ? ` (line ${lineMatch[1]})` : '';
        vscode.window.showWarningMessage(
            `Formatting skipped — Prettier could not parse the HTML${location}. ` +
            `This is usually caused by a missing or extra HTML tag. Fix the highlighted warnings first.`
        );
        return code;
    }

    // ── Step 3: Verify all placeholders survived Prettier ───────────────────

    for (const block of aspBlocks) {
        const needle =
            block.kind === 'inline' ? `ASPINLINE_${block.id}_END` :
            block.kind === 'midtag' ? `data-asp-${block.id}`       :
            block.id;

        if (!prettifiedCode.includes(needle)) {
            vscode.window.showWarningMessage(
                `Formatting skipped — an ASP block on line ${block.lineNumber + 1} was removed by Prettier. ` +
                `This usually happens when a <% %> block is in an unexpected position inside an HTML tag.`
            );
            return code;
        }
    }

    // ── Step 4: Format each ASP block's VBScript content ────────────────────
    // Process blocks sequentially so each normal block can thread its ending
    // indent level into the next, enabling cross-block continuity.

    const formattedBlocks: string[] = [];
    let   currentIndentLevel        = 0;

    for (const block of aspBlocks) {
        if (block.kind !== 'normal') {
            // Inline / midtag blocks: format but don't change the tracked level.
            const result = formatSingleAspBlock(block.code, aspSettings, '', currentIndentLevel);
            formattedBlocks.push(result.formatted);
            continue;
        }

        // For normal blocks, capture the HTML indent Prettier placed before
        // the placeholder comment so the VBScript formatter can use it when
        // htmlIndentMode is 'continuation'.
        const escapedId  = block.id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const indentMatch = prettifiedCode.match(new RegExp(`([ \\t]*)<!--${escapedId}-->`));
        const htmlIndent  = indentMatch ? indentMatch[1] : '';

        const result = formatSingleAspBlock(block.code, aspSettings, htmlIndent, currentIndentLevel);
        formattedBlocks.push(result.formatted);
        currentIndentLevel = result.endLevel;
    }

    // ── Step 5: Restore formatted ASP blocks into Prettier's output ─────────

    let restoredCode = prettifiedCode;

    for (let i = 0; i < aspBlocks.length; i++) {
        const block     = aspBlocks[i];
        const formatted = formattedBlocks[i];
        const escapedId = block.id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

        switch (block.kind) {

            case 'inline':
                // Replace the bare token directly — no indent adjustment needed.
                restoredCode = restoredCode.replace(
                    `ASPINLINE_${block.id}_END`,
                    formatted
                );
                break;

            case 'midtag':
                // Prettier may have normalised quotes/spacing around the attribute.
                restoredCode = restoredCode.replace(
                    new RegExp(`\\s*data-asp-${escapedId}\\s*=\\s*["']1["']`),
                    ` ${formatted}`
                );
                break;

            default: {
                // Pick up whatever indentation Prettier assigned to the comment
                // line, then prepend it to every non-empty line of the formatted
                // ASP block (the VBScript formatter itself handles internal
                // indentation relative to that base).
                const match = restoredCode.match(new RegExp(`([ \\t]*)<!--${escapedId}-->`));

                if (match) {
                    const htmlIndent    = match[1];
                    const indentedBlock = formatted
                        .split('\n')
                        .map(line => (line.trim() ? htmlIndent + line : line))
                        .join('\n');

                    restoredCode = restoredCode.replace(
                        new RegExp(`[ \\t]*<!--${escapedId}-->`),
                        indentedBlock
                    );
                } else {
                    // Safety fallback — the placeholder-lost check above should
                    // have already caught real deletions, but guard anyway.
                    restoredCode = restoredCode.replace(`<!--${block.id}-->`, formatted);
                }
                break;
            }
        }
    }

    return restoredCode;
}