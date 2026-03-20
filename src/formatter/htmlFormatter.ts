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
    // Strategy: scan backwards to find the nearest enclosing HTML attribute quote.
    //
    // HTML attributes use double-quotes `"`, while JS strings inside them use
    // single-quotes `'`.  When scanning backwards, any quote character (`"` or `'`)
    // that is immediately preceded by `=` (ignoring whitespace) is an HTML
    // attribute opening quote — meaning our ASP block is inside that attribute
    // value, so it is 'inline'.
    //
    // All other quotes are JS-string boundaries inside the attribute value.
    // We track them with a stack so we correctly skip their contents.
    //
    // Example (after Prettier wraps across lines):
    //   onclick="
    //       someFunc(
    //           '<%= val %>',   ← scanning backwards from here
    //       )
    //   "
    // Backwards path: ' (JS string close, push "'") → newline, spaces → " (not
    // matching top of stack which is "'", but preceded by "=" → return 'inline').

    let i = offset - 1;

    // Stack of quote chars for JS strings. Top = the innermost open JS string.
    const quoteStack: string[] = [];

    while (i >= 0) {
        const ch = code[i];

        // ── Check every quote for the HTML attribute open-quote pattern ──────
        // Regardless of nesting level, if a quote char is preceded by `=` it is
        // an HTML attribute boundary that encloses the ASP block.
        if (ch === '"' || ch === "'") {
            let j = i - 1;
            while (j >= 0 && (code[j] === ' ' || code[j] === '\t' || code[j] === '\n' || code[j] === '\r')) {
                j--;
            }
            if (j >= 0 && code[j] === '=') {
                // HTML attribute opening quote — the ASP block lives inside this value.
                return 'inline';
            }

            // Not an attribute open-quote. Treat as JS string boundary.
            if (quoteStack.length > 0 && quoteStack[quoteStack.length - 1] === ch) {
                // Matches the innermost open string — this is its opening quote, pop.
                quoteStack.pop();
            } else {
                // Closing quote of a JS string we haven't entered yet — push.
                quoteStack.push(ch);
            }
            i--; continue;
        }

        // ── Tag boundary characters (only relevant at top-level) ────────────
        // If we're nested inside a JS string, angle brackets are just content.
        if (quoteStack.length === 0) {
            if (ch === '>') {
                // If this > is part of a %> closing tag, skip the whole <% %> block.
                if (i > 0 && code[i - 1] === '%') {
                    i -= 2;
                    while (i >= 0) {
                        if (code[i] === '<' && i + 1 < code.length && code[i + 1] === '%') {
                            i--;
                            break;
                        }
                        i--;
                    }
                    continue;
                }
                return 'normal';
            }

            if (ch === '<') {
                const after = code.substring(i + 1, i + 3);
                if (after.startsWith('/'))        return 'normal';
                if (/^[a-zA-Z!?]/.test(after))   return 'midtag';
                return 'normal';
            }
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
                    // Check whether this %> is on a VBScript comment line.
                    // A VBScript comment starts with ' as the first non-whitespace
                    // character INSIDE the ASP block — not just anywhere on the
                    // HTML line.  We scan from the later of: the start of the
                    // current line, or the opening <% tag itself, so that a JS
                    // single-quote that precedes the ASP block on the same line
                    // (e.g.  '<%= cmpy %>'  ) cannot trip the comment check.
                    const lineBegin      = code.lastIndexOf('\n', end - 1) + 1;
                    const aspContentStart = pos + 2; // first char after <%
                    const scanFrom       = Math.max(lineBegin, aspContentStart);
                    const lineUpToClose  = code.slice(scanFrom, end).trimStart();
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
        });
    } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        const lineMatch = msg.match(/\((\d+):(\d+)\)/);
        const location  = lineMatch ? ` (line ${lineMatch[1]}, col ${lineMatch[2]})` : '';

        // ── Debug: log the masked code so we can see what Prettier choked on ──
        const channel = vscode.window.createOutputChannel('ASP Formatter Debug');
        channel.clear();
        channel.appendLine('=== Prettier parse error' + location + ' ===');
        channel.appendLine('Error: ' + msg);
        channel.appendLine('');
        channel.appendLine('=== Masked code sent to Prettier ===');
        channel.appendLine(maskedCode);
        channel.appendLine('');
        channel.appendLine('=== ASP blocks classified ===');
        for (const b of aspBlocks) {
            channel.appendLine(`  line ${b.lineNumber + 1}  kind=${b.kind}  ${b.code.slice(0, 60).replace(/\n/g, '\\n')}`);
        }
        channel.show(true);

        vscode.window.showWarningMessage(
            `Formatting skipped — Prettier could not parse the HTML${location}. ` +
            `Check the "ASP Formatter Debug" output channel to see the masked code.`
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

    const formattedBlocks:  string[] = [];
    const blockStartLevels: number[] = [];
    const blockHtmlIndents: string[] = [];
    let   currentIndentLevel         = 0;

    for (const block of aspBlocks) {
        if (block.kind !== 'normal') {
            // Inline / midtag blocks: format but don't change the tracked level.
            const result = formatSingleAspBlock(block.code, aspSettings, '', currentIndentLevel);
            formattedBlocks.push(result.formatted);
            blockStartLevels.push(-1);
            blockHtmlIndents.push('');
            continue;
        }

        // For normal blocks, capture the HTML indent Prettier placed before
        // the placeholder comment so the VBScript formatter can use it when
        // htmlIndentMode is 'continuation'.
        const escapedId  = block.id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const indentMatch = prettifiedCode.match(new RegExp(`([ \\t]*)<!--${escapedId}-->`));
        const htmlIndent  = indentMatch ? indentMatch[1] : '';

        blockStartLevels.push(currentIndentLevel);
        blockHtmlIndents.push(htmlIndent);

        const result = formatSingleAspBlock(block.code, aspSettings, htmlIndent, currentIndentLevel);
        formattedBlocks.push(result.formatted);
        currentIndentLevel = result.endLevel;
    }

    // ── Step 4b: Compute the shared tag-column for each linked group ─────────
    // In 'flat' mode, consecutive normal blocks that are part of the same
    // VBScript flow (startLevel > 0 for all but the first) share one logical
    // script.  Their <% / %> delimiters should all sit at the same column —
    // the HTML indent of the shallowest (first) block in the group — so the
    // tags form a consistent left margin regardless of HTML nesting depth.
    // A new group starts whenever a normal block has startLevel === 0.

    const blockTagIndents: string[] = new Array(aspBlocks.length).fill('');

    if (aspSettings.htmlIndentMode !== 'continuation') {
        let groupTagIndent = '';
        for (let i = 0; i < aspBlocks.length; i++) {
            if (aspBlocks[i].kind !== 'normal') continue;
            if (blockStartLevels[i] === 0) {
                // New group — this block's HTML indent becomes the shared tag column.
                groupTagIndent = blockHtmlIndents[i];
            }
            blockTagIndents[i] = groupTagIndent;
        }
    }

    // ── Step 5: Restore formatted ASP blocks into Prettier's output ─────────

    let restoredCode = prettifiedCode;

    for (let i = 0; i < aspBlocks.length; i++) {
        const block     = aspBlocks[i];
        const formatted = formattedBlocks[i];
        const escapedId = block.id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

        switch (block.kind) {

            case 'inline': {
                const inlineToken  = `ASPINLINE_${block.id}_END`;
                const escapedToken = inlineToken.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                const isExpression = block.code.trimStart().startsWith('<%=') ||
                                     block.code.trimStart().startsWith('<% =');

                if (isExpression) {
                    // <%= ... %> is an inline expression — part of the HTML content.
                    // Restore it exactly in place without touching any surrounding
                    // whitespace.  Prettier should not have wrapped it (we pass
                    // embeddedLanguageFormatting:'off'), but even if it did we keep
                    // the expression on one line because breaking it would introduce
                    // unwanted whitespace into the rendered HTML.
                    restoredCode = restoredCode.replace(inlineToken, formatted);
                } else {
                    // <% ... %> is a code block inside an attribute value.
                    // Prettier may have moved the token onto its own line — strip
                    // any newline + whitespace it injected before/after the token
                    // so the result doesn't break the attribute value open.
                    restoredCode = restoredCode.replace(
                        new RegExp(`(\\n[ \\t]*)?${escapedToken}([ \\t]*\\n)?`),
                        () => formatted
                    );
                }
                break;
            }

            case 'midtag':
                // Prettier may have normalised quotes/spacing around the attribute.
                restoredCode = restoredCode.replace(
                    new RegExp(`\\s*data-asp-${escapedId}\\s*=\\s*["']1["']`),
                    ` ${formatted}`
                );
                break;

            default: {
                const match = restoredCode.match(new RegExp(`([ \\t]*)<!--${escapedId}-->`));

                if (match) {
                    const lineIndent = match[1];

                    const placeholderIdx   = restoredCode.indexOf(`<!--${block.id}-->`);
                    const lineStart        = restoredCode.lastIndexOf('\n', placeholderIdx - 1) + 1;
                    const textBeforeOnLine = restoredCode.slice(lineStart, placeholderIdx);

                    // A block is "inline placed" when Prettier left non-whitespace
                    // content before the placeholder on the same line.
                    // EXCEPTION: if that content contains an unclosed quote the
                    // placeholder is sitting inside an attribute value
                    // (e.g. value="<!--ID-->" or onclick="...<!--ID-->...").
                    // In that case we must restore inline — no newlines inserted —
                    // because breaking the attribute value open corrupts the HTML.
                    const hasContentBefore = textBeforeOnLine.trimStart().length > 0;
                    const isInsideQuote    = hasContentBefore && (() => {
                        let inQ: string | null = null;
                        for (const ch of textBeforeOnLine) {
                            if (!inQ && (ch === '"' || ch === "'")) { inQ = ch; }
                            else if (inQ && ch === inQ)             { inQ = null; }
                        }
                        return inQ !== null; // still inside a quote → unclosed
                    })();
                    const isExpression     = block.code.trimStart().startsWith('<%=') ||
                                             block.code.trimStart().startsWith('<% =');
                    const isInlinePlaced   = hasContentBefore && !isInsideQuote && !isExpression;

                    if (isInlinePlaced && !aspSettings.aspTagsOnSameLine) {
                        // Block sits inline in tag text content (e.g. <td><!--ID--></td>).
                        // Expand it onto its own indented lines.
                        const indentUnit  = aspSettings.useTabs ? '\t' : ' '.repeat(aspSettings.indentSize);
                        const baseIndent  = textBeforeOnLine.match(/^([ \t]*)/)?.[1] ?? '';
                        const blockIndent = baseIndent + indentUnit;

                        // In flat mode use the group's shared tag column so this block's
                        // tags align with all sibling blocks in the same VBScript group.
                        const inlineTagIndent = aspSettings.htmlIndentMode === 'continuation'
                            ? ''
                            : blockTagIndents[i];

                        const indentedBlock = formatted
                            .split('\n')
                            .map(line => {
                                if (!line.trim()) return line;
                                const t = line.trim();
                                if (t === '<%' || t === '%>') return inlineTagIndent + t;
                                return aspSettings.htmlIndentMode === 'continuation'
                                    ? line
                                    : blockTagIndents[i] + line;
                            })
                            .join('\n');

                        restoredCode = restoredCode.replace(
                            new RegExp(`[ \\t]*<!--${escapedId}-->`),
                            `\n${indentedBlock}\n${baseIndent}`
                        );
                    } else {
                        // Expression sitting inline in tag text content (e.g. <td><%= val %></td>)
                        // — restore it exactly as-is with no indentation applied.
                        if (isExpression && hasContentBefore && !isInsideQuote) {
                            restoredCode = restoredCode.replace(
                                new RegExp(`[ \\t]*<!--${escapedId}-->`),
                                formatted.trim()
                            );
                            break;
                        }

                        // Block is either standalone on its own line, or inside a
                        // quoted attribute value — restore with correct tag indentation.
                        //
                        // How <% and %> tag lines are indented depends on htmlIndentMode:
                        //
                        // 'flat' mode  — aspFormatter starts VBScript at level 0, so
                        //   content lines have only VBScript indent (e.g. "    If ...").
                        //   The <% / %> tags should sit at the HTML placeholder indent
                        //   (lineIndent) so they visually belong to the HTML structure,
                        //   and content is indented further in from there.
                        //   e.g.  <div>\n  <% ← lineIndent, content at lineIndent+vbsIndent
                        //
                        // 'continuation' mode — aspFormatter already adds HTML depth to
                        //   content indent, so content is fully self-contained.
                        //   The <% / %> tags go to column 0 to avoid double-indenting.
                        //   e.g.  <%\n        If ... (HTML+VBScript indent already baked in)
                        // 'flat' mode: use the group's shared tag column (the HTML indent
                        // of the first/shallowest block in this VBScript group) so all
                        // <% / %> tags in the group align at the same column.
                        const tagIndent = aspSettings.htmlIndentMode === 'continuation'
                            ? ''                   // tags at col 0; content has full indent
                            : blockTagIndents[i];  // shared group column

                        const indentedBlock = formatted
                            .split('\n')
                            .map(line => {
                                if (!line.trim()) return line;
                                const t = line.trim();
                                if (t === '<%' || t === '%>') return tagIndent + t;
                                // Content lines: in flat mode add the group tag indent on
                                // top of the VBScript indent aspFormatter produced.
                                // In continuation mode keep as-is (full indent baked in).
                                return aspSettings.htmlIndentMode === 'continuation'
                                    ? line
                                    : blockTagIndents[i] + line;
                            })
                            .join('\n');

                        restoredCode = restoredCode.replace(
                            new RegExp(`[ \\t]*<!--${escapedId}-->`),
                            indentedBlock
                        );
                    }
                } else {
                    restoredCode = restoredCode.replace(`<!--${block.id}-->`, formatted);
                }
                break;
            }
        }
    }

    // ── Step 6: Fix broken whitespace-sensitive tags (e.g. <textarea>) ───────
    // When ASPINLINE tokens are long, Prettier wraps the closing `>` of the
    // opening tag onto its own line, and separately breaks the closing tag:
    //
    //   <textarea ...attrs...>
    //   <%= val %></textarea
    //                       >
    //
    // This is invalid for whitespace-sensitive tags because any whitespace
    // between `>` and the content becomes part of the rendered text.
    // Collapse both splits so the result is:
    //
    //   <textarea ...attrs...><%= val %></textarea>
    //
    // Pattern:
    //   (>)           — the closing > of the opening tag (already on its own line or inline)
    //   \n[ \t]*      — newline + any indentation Prettier added before the content
    //   ([^\n]+?)     — the inline content (single line, non-greedy)
    //   (<\/\w+)      — start of the closing tag (e.g. </textarea)
    //   \n[ \t]*      — newline + whitespace before the stray >
    //   (>)           — the stray > that closes the closing tag
    restoredCode = restoredCode.replace(
        /(>)\n[ \t]*([^\n]+?)(<\/\w+)\n[ \t]*(>)/g,
        '$1$2$3$4'
    );

    return restoredCode;
}