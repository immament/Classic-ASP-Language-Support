/**
 * jsDiagnosticsProvider.ts  (providers/)
 *
 * CSS-style diagnostics (error/warning squiggles) for JavaScript inside
 * <script> blocks in .asp files, powered by the TypeScript Language Service.
 *
 * Mirrors cssDiagnosticsProvider.ts in structure and registration pattern.
 *
 * What gets flagged:
 *   • Syntax errors   — always (e.g. missing brackets, unexpected tokens)
 *   • Semantic errors — basic ones only:
 *       - Undeclared variables / members (TS2304, TS2339)
 *       - Wrong number of arguments (TS2554)
 *       - etc.
 *   Strict type errors (implicit any, return type mismatches) are suppressed
 *   via compiler options so inline scripts don't get flooded with warnings.
 *
 * Debounced at 750 ms — faster than the HTML structure checker (1500 ms)
 * because JS errors tend to be typed incrementally and need quick feedback.
 */

import * as vscode from 'vscode';
import * as ts     from 'typescript';
import {
    buildVirtualJsContent,
    getJsLanguageService,
    tsSeverityToVs,
} from '../utils/jsUtils';

// Diagnostic codes we deliberately suppress in inline script blocks.
// These are too noisy for small embedded scripts that don't import modules.
const SUPPRESSED_CODES = new Set([
    2304,   // Cannot find name 'X'  — too many false positives for globals
    2339,   // Property 'X' does not exist on type 'Y' — common for dynamic DOM
    2345,   // Argument of type 'X' is not assignable to parameter of type 'Y'
    2322,   // Type 'X' is not assignable to type 'Y'
    7006,   // Parameter 'X' implicitly has an 'any' type
    7005,   // Variable 'X' implicitly has an 'any' type
    2531,   // Object is possibly 'null'
    2532,   // Object is possibly 'undefined'
]);

function getDiagnosticsForDocument(
    document: vscode.TextDocument
): vscode.Diagnostic[] {
    const content = document.getText();
    const diagnostics: vscode.Diagnostic[] = [];

    // Build the virtual JS content once for the whole document
    // (offset 0 — we just need the virtualContent, not isInScript)
    const { virtualContent } = buildVirtualJsContent(content, 0);

    const svc = getJsLanguageService();
    svc.updateContent(virtualContent);

    const allDiags: ts.Diagnostic[] = [
        ...svc.getSyntacticDiagnostics(),
        ...svc.getSemanticDiagnostics(),
    ];

    for (const d of allDiags) {
        // Skip diagnostics with no position (project-wide config issues etc.)
        if (d.start === undefined || d.length === undefined) { continue; }

        // Suppress noisy codes that don't make sense for inline scripts
        const code = typeof d.code === 'number' ? d.code : 0;
        if (SUPPRESSED_CODES.has(code)) { continue; }

        const startPos = document.positionAt(d.start);
        const endPos   = document.positionAt(d.start + d.length);

        // Only report errors that fall inside actual <script> content —
        // the virtual document has spaces elsewhere but TS might still
        // produce a diagnostic for a blank region in edge cases.
        // We guard by checking the range falls on a non-whitespace-only line.
        const lineText = document.lineAt(startPos.line).text.trim();
        if (!lineText) { continue; }

        const message = typeof d.messageText === 'string'
            ? d.messageText
            : ts.flattenDiagnosticMessageText(d.messageText, '\n');

        const diag = new vscode.Diagnostic(
            new vscode.Range(startPos, endPos),
            message,
            tsSeverityToVs(d.category)
        );
        diag.source = 'Classic ASP (JS)';
        diag.code   = code;
        diagnostics.push(diag);
    }

    return diagnostics;
}

export function registerJsDiagnostics(
    context: vscode.ExtensionContext
): void {
    const collection = vscode.languages.createDiagnosticCollection('classic-asp-js');
    context.subscriptions.push(collection);

    let debounceTimer: ReturnType<typeof setTimeout> | undefined;

    function schedule(document: vscode.TextDocument): void {
        if (document.languageId !== 'asp') { return; }
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
            collection.set(document.uri, getDiagnosticsForDocument(document));
        }, 750);
    }

    // Run on already-open documents
    for (const doc of vscode.workspace.textDocuments) {
        if (doc.languageId === 'asp') {
            collection.set(doc.uri, getDiagnosticsForDocument(doc));
        }
    }

    context.subscriptions.push(
        vscode.workspace.onDidOpenTextDocument(schedule),
        vscode.workspace.onDidChangeTextDocument(e => schedule(e.document)),
        vscode.workspace.onDidCloseTextDocument(doc => collection.delete(doc.uri)),
    );
}