/**
 * jsHoverProvider.ts  (providers/)
 *
 * Shows TypeScript Quick Info (type + JSDoc) when hovering over a symbol
 * inside a <script> block in a .asp file.
 *
 * Mirrors CssHoverProvider in structure and registration pattern.
 */

import * as vscode from 'vscode';
import {
    buildVirtualJsContent,
    getJsLanguageService,
    isInJsZone,
} from '../utils/jsUtils';

export class JsHoverProvider implements vscode.HoverProvider {

    provideHover(
        document: vscode.TextDocument,
        position: vscode.Position,
        token:    vscode.CancellationToken
    ): vscode.ProviderResult<vscode.Hover> {

        if (!isInJsZone(document, position)) { return undefined; }

        const offset  = document.offsetAt(position);
        const content = document.getText();
        const { virtualContent, isInScript } = buildVirtualJsContent(content, offset);
        if (!isInScript || token.isCancellationRequested) { return undefined; }

        const svc = getJsLanguageService();
        svc.updateContent(virtualContent);

        const info = svc.getQuickInfo(offset);
        if (!info || token.isCancellationRequested) { return undefined; }

        const displayText = info.displayParts?.map(p => p.text).join('') ?? '';
        const docsText    = info.documentation?.map(p => p.text).join('') ?? '';

        if (!displayText && !docsText) { return undefined; }

        const md = new vscode.MarkdownString();
        if (displayText) { md.appendCodeblock(displayText, 'typescript'); }
        if (docsText)    { md.appendMarkdown('\n\n' + docsText); }

        // Map the TS text span back to a real VS Code range so the hover
        // highlights the correct word in the editor.
        let range: vscode.Range | undefined;
        if (info.textSpan) {
            range = new vscode.Range(
                document.positionAt(info.textSpan.start),
                document.positionAt(info.textSpan.start + info.textSpan.length)
            );
        }

        return new vscode.Hover(md, range);
    }
}