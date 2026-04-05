/**
 * jsCompletionProvider.ts  (providers/)
 *
 * TypeScript Language Service completions for <script> blocks.
 *
 * Resolution data is stored on item.data (cast through `any` since older
 * @types/vscode versions don't declare this field publicly) rather than a
 * WeakMap, so it survives VS Code's internal serialize/deserialize cycle
 * between provideCompletionItems and resolveCompletionItem.
 *
 * Trigger characters are limited to '.' and '(' — VS Code's built-in
 * word-based filter handles the list once it is returned with isIncomplete:false,
 * so there is no need to register every letter.
 */

import * as vscode from 'vscode';
import {
    buildVirtualJsContent,
    getJsLanguageService,
    isInJsZone,
    tsKindToVsKind,
} from '../utils/jsUtils';

interface ItemData { name: string; offset: number; source?: string }

export class JsCompletionProvider implements vscode.CompletionItemProvider {

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
            item.sortText   = '0' + (entry.sortText ?? entry.name);
            item.filterText = entry.name;

            if (entry.insertText) {
                item.insertText = entry.isSnippet
                    ? new vscode.SnippetString(entry.insertText)
                    : entry.insertText;
            }

            if (item.kind === vscode.CompletionItemKind.Function ||
                item.kind === vscode.CompletionItemKind.Method) {
                item.commitCharacters = ['('];
            }

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (item as any).data = { name: entry.name, offset, source: entry.source } satisfies ItemData;

            return item;
        });

        return new vscode.CompletionList(items, false);
    }

    resolveCompletionItem(
        item:  vscode.CompletionItem,
        token: vscode.CancellationToken
    ): vscode.ProviderResult<vscode.CompletionItem> {

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const data = (item as any).data as ItemData | undefined;
        if (!data || token.isCancellationRequested) { return item; }

        const details = getJsLanguageService().getCompletionDetails(data.name, data.offset, data.source);
        if (!details || token.isCancellationRequested) { return item; }

        const displayText = details.displayParts?.map(p => p.text).join('') ?? '';
        const docsText    = details.documentation?.map(p => p.text).join('') ?? '';
        const tagsText    = details.tags?.map(tag => {
            const body = tag.text?.map(p => p.text).join('') ?? '';
            return body ? `*@${tag.name}* — ${body}` : `*@${tag.name}*`;
        }).join('\n\n') ?? '';

        if (displayText) { item.detail = displayText; }

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