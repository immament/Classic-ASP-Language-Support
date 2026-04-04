/**
 * jsUtils.ts  (utils/)
 *
 * Embedded JavaScript support for .asp files — TypeScript Compiler API edition.
 *
 * Architecture mirrors cssUtils.ts exactly:
 *   • buildVirtualJsContent()  — extracts <script> regions, blanks everything
 *                                else to preserve character offsets
 *   • JsLanguageService        — thin wrapper around ts.createLanguageService
 *                                (one singleton, reused across all providers)
 *
 * No separate language server process.  No new npm runtime dependencies beyond
 * the `typescript` package that every extension project already has.
 *
 * Compiler settings (browser-oriented):
 *   target  ES2020   allowJs true   strict false   noEmit true
 *   lib     lib.dom.d.ts + lib.es2020.d.ts   → full window/document/fetch types
 */

import * as path from 'path';
import * as ts   from 'typescript';
import * as vscode from 'vscode';
import { getZone } from './aspUtils';

// ─────────────────────────────────────────────────────────────────────────────
// Virtual file name
// Must end in .js so TypeScript uses loose JS checking (no type errors for
// untyped variables, etc.).
// ─────────────────────────────────────────────────────────────────────────────
const VIRTUAL_FILENAME = 'asp-embedded.js';

// ─────────────────────────────────────────────────────────────────────────────
// buildVirtualJsContent
//
// Returns the full document text with every character outside <script> blocks
// replaced by a space (newlines preserved so line numbers stay accurate).
// ASP blocks (<% … %>) that survived into the JS zone are also blanked so the
// TS parser never sees them.
//
// Offset alignment is exact: virtualContent[n] corresponds to the same
// character position as document.getText()[n].
// ─────────────────────────────────────────────────────────────────────────────
export interface VirtualJsResult {
    virtualContent: string;
    isInScript:     boolean;    // true when `offset` falls inside a <script> block
}

export function buildVirtualJsContent(
    content: string,
    offset:  number
): VirtualJsResult {
    // Collect [start, end) pairs for every JS <script> block content region.
    const jsRanges: Array<{ start: number; end: number }> = [];

    const scriptOpenRe = /<script(\s[^>]*)?>/gi;
    let m: RegExpExecArray | null;

    while ((m = scriptOpenRe.exec(content)) !== null) {
        const attrs    = m[1] ?? '';
        const tagEnd   = m.index + m[0].length;

        // Skip non-JS script types (text/template, text/x-handlebars, etc.)
        const typeMatch = attrs.match(/\btype\s*=\s*["']([^"']+)["']/i);
        if (typeMatch && !/javascript|module/i.test(typeMatch[1])) { continue; }

        // Also skip VBScript blocks
        if (/\blanguage\s*=\s*["']vbscript["']/i.test(attrs)) { continue; }

        // Find matching </script>
        const closeIdx = content.search(new RegExp('<\\/script\\s*>', 'i'));
        // Use indexOf for speed — we know the search starts at tagEnd
        const closeTagRe = /<\/script\s*>/i;
        closeTagRe.lastIndex = 0;
        const rest = content.slice(tagEnd);
        const closeM = rest.search(closeTagRe);
        const end = closeM === -1 ? content.length : tagEnd + closeM;

        jsRanges.push({ start: tagEnd, end });
    }

    // Check whether offset is inside a JS region
    const isInScript = jsRanges.some(r => offset >= r.start && offset <= r.end);

    // Build virtual content: blank all non-JS regions character by character
    const chars = content.split('');
    for (let i = 0; i < chars.length; i++) {
        const inJs = jsRanges.some(r => i >= r.start && i < r.end);
        if (!inJs && chars[i] !== '\n') { chars[i] = ' '; }
    }
    let virtualContent = chars.join('');

    // Blank ASP blocks that leaked into the JS zone
    virtualContent = virtualContent.replace(/<%[\s\S]*?%>/g, m =>
        m.replace(/[^\n]/g, ' ')
    );

    return { virtualContent, isInScript };
}

// ─────────────────────────────────────────────────────────────────────────────
// JsLanguageService
// ─────────────────────────────────────────────────────────────────────────────
export class JsLanguageService implements vscode.Disposable {
    private readonly _service: ts.LanguageService;
    private          _content: string = '';
    private          _version: number = 0;

    constructor() {
        const compilerOptions: ts.CompilerOptions = {
            target:                ts.ScriptTarget.ES2020,
            lib:                   ['lib.dom.d.ts', 'lib.es2020.d.ts'],
            allowJs:               true,
            noEmit:                true,
            strict:                false,
            checkJs:               false,
            noSemanticValidation:  true,
        } as ts.CompilerOptions;

        // Locate TypeScript's own bundled lib directory so lib.dom.d.ts resolves.
        const libDir = path.dirname(ts.getDefaultLibFilePath(compilerOptions));

        const host: ts.LanguageServiceHost = {
            getScriptFileNames:  () => [VIRTUAL_FILENAME],
            getScriptVersion:    (f) => f === VIRTUAL_FILENAME ? String(this._version) : '0',
            getScriptSnapshot:   (f) => {
                if (f === VIRTUAL_FILENAME) {
                    return ts.ScriptSnapshot.fromString(this._content);
                }
                const text = ts.sys.readFile(f);
                return text !== undefined ? ts.ScriptSnapshot.fromString(text) : undefined;
            },
            getCompilationSettings: () => compilerOptions,
            getCurrentDirectory:    () => libDir,
            getDefaultLibFileName:  (opts) => ts.getDefaultLibFilePath(opts),
            fileExists:  (f) => f === VIRTUAL_FILENAME || ts.sys.fileExists(f),
            readFile:    (f) => f === VIRTUAL_FILENAME ? this._content : ts.sys.readFile(f),
            readDirectory:   ts.sys.readDirectory.bind(ts.sys),
            directoryExists: ts.sys.directoryExists.bind(ts.sys),
            getDirectories:  ts.sys.getDirectories.bind(ts.sys),
        };

        this._service = ts.createLanguageService(host, ts.createDocumentRegistry());
    }

    /** Call before every request with the current virtual JS content. */
    updateContent(content: string): void {
        this._content = content;
        this._version++;
    }

    getCompletions(offset: number, trigger?: string): ts.CompletionInfo | undefined {
        try {
            return this._service.getCompletionsAtPosition(VIRTUAL_FILENAME, offset, {
                triggerCharacter: trigger as ts.CompletionsTriggerCharacter | undefined,
            }) ?? undefined;
        } catch { return undefined; }
    }

    getCompletionDetails(
        name: string, offset: number, source?: string
    ): ts.CompletionEntryDetails | undefined {
        try {
            return this._service.getCompletionEntryDetails(
                VIRTUAL_FILENAME, offset, name,
                undefined, source, undefined, undefined
            ) ?? undefined;
        } catch { return undefined; }
    }

    getQuickInfo(offset: number): ts.QuickInfo | undefined {
        try {
            return this._service.getQuickInfoAtPosition(VIRTUAL_FILENAME, offset) ?? undefined;
        } catch { return undefined; }
    }

    getSignatureHelp(offset: number): ts.SignatureHelpItems | undefined {
        try {
            return this._service.getSignatureHelpItems(
                VIRTUAL_FILENAME, offset, undefined
            ) ?? undefined;
        } catch { return undefined; }
    }

    dispose(): void {
        this._service.dispose();
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Singleton — one service shared across all providers
// ─────────────────────────────────────────────────────────────────────────────
let _service: JsLanguageService | undefined;

export function getJsLanguageService(): JsLanguageService {
    if (!_service) { _service = new JsLanguageService(); }
    return _service;
}

export function disposeJsLanguageService(): void {
    _service?.dispose();
    _service = undefined;
}

// ─────────────────────────────────────────────────────────────────────────────
// Zone helper
// ─────────────────────────────────────────────────────────────────────────────
export function isInJsZone(
    document: vscode.TextDocument,
    position: vscode.Position
): boolean {
    return getZone(document.getText(), document.offsetAt(position)) === 'js';
}

// ─────────────────────────────────────────────────────────────────────────────
// ts.ScriptElementKind → vscode.CompletionItemKind
// ─────────────────────────────────────────────────────────────────────────────
export function tsKindToVsKind(kind: string): vscode.CompletionItemKind {
    switch (kind) {
        case ts.ScriptElementKind.functionElement:
        case ts.ScriptElementKind.localFunctionElement:
            return vscode.CompletionItemKind.Function;
        case ts.ScriptElementKind.memberFunctionElement:
        case ts.ScriptElementKind.callSignatureElement:
        case ts.ScriptElementKind.constructSignatureElement:
            return vscode.CompletionItemKind.Method;
        case ts.ScriptElementKind.variableElement:
        case ts.ScriptElementKind.localVariableElement:
        case ts.ScriptElementKind.letElement:
        case ts.ScriptElementKind.constElement:
            return vscode.CompletionItemKind.Variable;
        case ts.ScriptElementKind.classElement:
        case ts.ScriptElementKind.localClassElement:
            return vscode.CompletionItemKind.Class;
        case ts.ScriptElementKind.interfaceElement:
            return vscode.CompletionItemKind.Interface;
        case ts.ScriptElementKind.enumElement:
            return vscode.CompletionItemKind.Enum;
        case ts.ScriptElementKind.enumMemberElement:
            return vscode.CompletionItemKind.EnumMember;
        case ts.ScriptElementKind.moduleElement:
        case ts.ScriptElementKind.externalModuleName:
            return vscode.CompletionItemKind.Module;
        case ts.ScriptElementKind.memberVariableElement:
        case ts.ScriptElementKind.memberGetAccessorElement:
        case ts.ScriptElementKind.memberSetAccessorElement:
            return vscode.CompletionItemKind.Field;
        case ts.ScriptElementKind.typeElement:
        case ts.ScriptElementKind.typeParameterElement:
            return vscode.CompletionItemKind.TypeParameter;
        case ts.ScriptElementKind.keyword:
            return vscode.CompletionItemKind.Keyword;
        case ts.ScriptElementKind.string:
            return vscode.CompletionItemKind.Value;
        default:
            return vscode.CompletionItemKind.Property;
    }
}