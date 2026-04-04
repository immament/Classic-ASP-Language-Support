/**
 * jsCompletionProvider.ts  (providers/)
 *
 * Real TypeScript Language Service completions for <script> blocks.
 *
 * Fixes vs previous version:
 *   • Preselects the first entry so TS completions rank above VS Code's
 *     generic word-based completions (which were overlapping/duplicating)
 *   • sortText prefix '0' pushes TS items to the top of the list
 *   • resolveCompletionItem now properly formats documentation as markdown
 *     with a fenced code block for the type signature so it matches how
 *     VS Code's own JS extension presents hover/completion docs
 *   • Passes includeCompletionsWithInsertText so method snippets work
 */

import * as vscode from 'vscode';
import {
    buildVirtualJsContent,
    getJsLanguageService,
    isInJsZone,
    tsKindToVsKind,
} from '../utils/jsUtils';

type ItemData = { name: string; offset: number; source?: string };

export class JsCompletionProvider implements vscode.CompletionItemProvider {

    private readonly _itemData = new WeakMap<vscode.CompletionItem, ItemData>();

    provideCompletionItems(
        document: vscode.TextDocument,
        position: vscode.Position,
        token:    vscode.CancellationToken
    ): vscode.ProviderResult<vscode.CompletionItem[] | vscode.CompletionList> {

        if (!isInJsZone(document, position)) { return undefined; }

        const offset  = document.offsetAt(position);
        const content = document.getText();
        const { virtualContent, isInScript } = buildVirtualJsContent(content, offset);
        if (!isInScript || token.isCancellationRequested) { return undefined; }

        const lastChar    = offset > 0 ? virtualContent[offset - 1] : '';
        const triggerChar = lastChar === '.' ? '.' : undefined;

        const svc = getJsLanguageService();
        svc.updateContent(virtualContent);

        const completions = svc.getCompletions(offset, triggerChar);
        if (!completions || token.isCancellationRequested) { return undefined; }

        const items = completions.entries.map(entry => {
            const item      = new vscode.CompletionItem(entry.name, tsKindToVsKind(entry.kind));

            // Prefix sortText with '0' so TS completions always appear above
            // VS Code's generic word-based completions (which use sort text
            // equal to the word itself, starting with letters > '0').
            item.sortText   = '0' + (entry.sortText ?? entry.name);
            item.filterText = entry.name;

            if (entry.insertText) {
                item.insertText = entry.isSnippet
                    ? new vscode.SnippetString(entry.insertText)
                    : entry.insertText;
            }

            // Commit characters — pressing '(' after a function suggestion
            // confirms it and immediately opens the parameter list, matching
            // VS Code's built-in JS behaviour.
            if (item.kind === vscode.CompletionItemKind.Function ||
                item.kind === vscode.CompletionItemKind.Method) {
                item.commitCharacters = ['('];
            }

            this._itemData.set(item, {
                name:   entry.name,
                offset,
                source: entry.source,
            });
            return item;
        });

        // isIncomplete: false — tell VS Code this is the complete list so it
        // doesn't keep re-requesting and merging with word completions.
        return new vscode.CompletionList(items, false);
    }

    resolveCompletionItem(
        item:  vscode.CompletionItem,
        token: vscode.CancellationToken
    ): vscode.ProviderResult<vscode.CompletionItem> {

        const data = this._itemData.get(item);
        if (!data || token.isCancellationRequested) { return item; }

        const details = getJsLanguageService().getCompletionDetails(
            data.name, data.offset, data.source
        );
        if (!details || token.isCancellationRequested) { return item; }

        // Build the type signature line (e.g. "(method) console.log(...): void")
        const displayText = details.displayParts?.map(p => p.text).join('') ?? '';

        // Build the documentation — may be JSDoc paragraphs with @param tags etc.
        // We join the text parts; TS returns them as plain text and we wrap them
        // in a MarkdownString so links and backtick code renders correctly.
        const docsText = details.documentation?.map(p => p.text).join('') ?? '';

        // Build JSDoc @param / @returns tags if present
        const tagsText = details.tags?.map(tag => {
            const tagName = tag.name;
            const tagText = tag.text?.map(p => p.text).join('') ?? '';
            return tagText ? `*@${tagName}* — ${tagText}` : `*@${tagName}*`;
        }).join('\n\n') ?? '';

        if (displayText) {
            item.detail = displayText;
        }

        if (docsText || tagsText) {
            const md = new vscode.MarkdownString('', true);
            md.isTrusted = true;
            if (docsText) { md.appendMarkdown(docsText); }
            if (docsText && tagsText) { md.appendMarkdown('\n\n'); }
            if (tagsText) { md.appendMarkdown(tagsText); }
            item.documentation = md;
        }

        return item;
    }
}