/**
 * jsCompletionProvider.ts  (providers/)
 *
 * Replaces the old hand-crafted keyword/method list completion with real
 * TypeScript Language Service completions — the same engine that powers
 * VS Code's own JavaScript IntelliSense.
 *
 * Provides:
 *   • Member-access completions:  document.|   console.|   window.|  etc.
 *   • Global completions:         fetch(   addEventListener(   etc.
 *   • User-declared symbols:      functions / variables defined in <script>
 *   • Full DOM / ES2020 types via lib.dom.d.ts + lib.es2020.d.ts
 *   • resolveCompletionItem — shows JSDoc on the selected item
 */

import * as vscode from 'vscode';
import * as ts     from 'typescript';
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

        // Detect trigger character from the virtual content so dot-completions
        // always fire even when the user did not trigger via the trigger list.
        const lastChar    = offset > 0 ? virtualContent[offset - 1] : '';
        const triggerChar = lastChar === '.' ? '.' : undefined;

        const svc = getJsLanguageService();
        svc.updateContent(virtualContent);

        const completions = svc.getCompletions(offset, triggerChar);
        if (!completions || token.isCancellationRequested) { return undefined; }

        const items = completions.entries.map(entry => {
            const item      = new vscode.CompletionItem(entry.name, tsKindToVsKind(entry.kind));
            item.sortText   = entry.sortText;
            item.filterText = entry.name;

            if (entry.insertText) {
                item.insertText = entry.isSnippet
                    ? new vscode.SnippetString(entry.insertText)
                    : entry.insertText;
            }

            this._itemData.set(item, {
                name:   entry.name,
                offset,
                source: entry.source,
            });
            return item;
        });

        return new vscode.CompletionList(items, completions.isIncomplete);
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

        const displayText = details.displayParts?.map(p => p.text).join('') ?? '';
        const docsText    = details.documentation?.map(p => p.text).join('') ?? '';

        if (displayText) { item.detail        = displayText; }
        if (docsText)    { item.documentation = new vscode.MarkdownString(docsText); }

        return item;
    }
}