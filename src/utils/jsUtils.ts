/**
 * jsUtils.ts  (utils/)
 *
 * Embedded JavaScript support for .asp files — TypeScript Compiler API edition.
 *
 * Fixes in this version:
 *   • buildVirtualJsContent — correct forward-only </script> search (was
 *     always searching from position 0, breaking multi-script files and
 *     causing textarea.addEventListener etc. to be missing)
 *   • CompilerOptions — added types:[] to block @types/node leaking in
 *     (was causing console to show Node.js docs instead of browser docs)
 *   • Lib resolution — explicit lib file list resolved via
 *     ts.getDefaultLibFilePath so DOM types always load correctly
 */

import * as path from 'path';
import * as ts   from 'typescript';
import * as vscode from 'vscode';
import { getZone } from './aspUtils';

// ─────────────────────────────────────────────────────────────────────────────
// Virtual file — must end in .js for loose JS checking
// ─────────────────────────────────────────────────────────────────────────────
export const VIRTUAL_FILENAME = 'asp-embedded.js';

// ─────────────────────────────────────────────────────────────────────────────
// buildVirtualJsContent
//
// Scans forward through the document locating every JS <script>…</script>
// block.  All characters outside those blocks are replaced with spaces
// (newlines preserved) so offset positions stay exact.
// ASP blocks (<% … %>) inside script zones are also blanked.
//
// BUG FIXED: previous version called content.search(re) which always
// searched from index 0.  Now we search forward from tagEnd each time.
// ─────────────────────────────────────────────────────────────────────────────
export interface VirtualJsResult {
    virtualContent: string;
    isInScript:     boolean;
}

export function buildVirtualJsContent(
    content: string,
    offset:  number
): VirtualJsResult {
    const jsRanges: Array<{ start: number; end: number }> = [];

    // Walk through the document finding all <script> openers
    const scriptOpenRe = /<script(\s[^>]*)?>/gi;
    let m: RegExpExecArray | null;

    while ((m = scriptOpenRe.exec(content)) !== null) {
        const attrs  = m[1] ?? '';
        const tagEnd = m.index + m[0].length;

        // Skip non-JS script types
        const typeMatch = attrs.match(/\btype\s*=\s*["']([^"']+)["']/i);
        if (typeMatch && !/javascript|module/i.test(typeMatch[1])) { continue; }
        if (/\blanguage\s*=\s*["']vbscript["']/i.test(attrs)) { continue; }

        // ── FIX: search for </script> FORWARD from tagEnd, not from 0 ──
        const closeRe  = /<\/script\s*>/i;
        const rest     = content.slice(tagEnd);
        const closeIdx = rest.search(closeRe);
        const end      = closeIdx === -1 ? content.length : tagEnd + closeIdx;

        jsRanges.push({ start: tagEnd, end });
        // Advance the outer regex past this block so it doesn't re-match
        // content inside the script body as a nested <script> opener.
        scriptOpenRe.lastIndex = end;
    }

    const isInScript = jsRanges.some(r => offset >= r.start && offset <= r.end);

    // Build virtual content
    const chars = content.split('');
    for (let i = 0; i < chars.length; i++) {
        const inJs = jsRanges.some(r => i >= r.start && i < r.end);
        if (!inJs && chars[i] !== '\n') { chars[i] = ' '; }
    }
    let virtualContent = chars.join('');

    // Blank any ASP blocks inside the JS zones
    virtualContent = virtualContent.replace(/<%[\s\S]*?%>/g,
        match => match.replace(/[^\n]/g, ' ')
    );

    return { virtualContent, isInScript };
}

// ─────────────────────────────────────────────────────────────────────────────
// Compiler options — browser-oriented, Node types explicitly blocked
// ─────────────────────────────────────────────────────────────────────────────
function makeBrowserCompilerOptions(): ts.CompilerOptions {
    return {
        target:      ts.ScriptTarget.ES2020,
        // Explicitly list libs — TypeScript resolves these as filenames
        // relative to its own lib directory.
        lib:         [
            'lib.es2020.d.ts',
            'lib.dom.d.ts',
            'lib.dom.iterable.d.ts',
        ],
        allowJs:     true,
        checkJs:     true,       // enables syntactic + basic semantic errors
        noEmit:      true,
        strict:      false,
        // ── IMPORTANT: block @types/node and any other ambient types ──
        // Without this, if the user's project has @types/node installed,
        // TypeScript picks up Node.js typings and console becomes the
        // Node.js Console (showing util.format docs, etc.).
        types:       [],
        // Keep errors manageable — don't flag missing return types etc.
        noImplicitAny:           false,
        noImplicitReturns:       false,
        noUnusedLocals:          false,
        noUnusedParameters:      false,
        strictNullChecks:        false,
        strictFunctionTypes:     false,
        strictPropertyInitialization: false,
    };
}

// ─────────────────────────────────────────────────────────────────────────────
// JsLanguageService
// ─────────────────────────────────────────────────────────────────────────────
export class JsLanguageService implements vscode.Disposable {
    private readonly _service:         ts.LanguageService;
    private readonly _compilerOptions: ts.CompilerOptions;
    private          _content:         string = '';
    private          _version:         number = 0;

    constructor() {
        this._compilerOptions = makeBrowserCompilerOptions();

        // Point TypeScript at its own lib directory so it can find
        // lib.dom.d.ts, lib.es2020.d.ts, etc. from disk.
        const libDir = path.dirname(
            ts.getDefaultLibFilePath(this._compilerOptions)
        );

        const self = this;
        const host: ts.LanguageServiceHost = {
            getScriptFileNames:     () => [VIRTUAL_FILENAME],
            getScriptVersion:       (f) => f === VIRTUAL_FILENAME ? String(self._version) : '0',
            getScriptSnapshot:      (f) => {
                if (f === VIRTUAL_FILENAME) {
                    return ts.ScriptSnapshot.fromString(self._content);
                }
                const text = ts.sys.readFile(f);
                return text !== undefined ? ts.ScriptSnapshot.fromString(text) : undefined;
            },
            getCompilationSettings: () => self._compilerOptions,
            getCurrentDirectory:    () => libDir,
            getDefaultLibFileName:  (opts) => ts.getDefaultLibFilePath(opts),
            fileExists:             (f) => f === VIRTUAL_FILENAME || ts.sys.fileExists(f),
            readFile:               (f) => f === VIRTUAL_FILENAME ? self._content : ts.sys.readFile(f),
            readDirectory:          ts.sys.readDirectory.bind(ts.sys),
            directoryExists:        ts.sys.directoryExists.bind(ts.sys),
            getDirectories:         ts.sys.getDirectories.bind(ts.sys),
        };

        this._service = ts.createLanguageService(host, ts.createDocumentRegistry());
    }

    updateContent(content: string): void {
        this._content = content;
        this._version++;
    }

    getCompletions(offset: number, trigger?: string): ts.CompletionInfo | undefined {
        try {
            return this._service.getCompletionsAtPosition(VIRTUAL_FILENAME, offset, {
                triggerCharacter: trigger as ts.CompletionsTriggerCharacter | undefined,
                includeCompletionsWithInsertText:      true,
                includeCompletionsForModuleExports:    false,
                includeAutomaticOptionalChainCompletions: true,
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

    getSyntacticDiagnostics(): ts.DiagnosticWithLocation[] {
        try {
            return this._service.getSyntacticDiagnostics(VIRTUAL_FILENAME) ?? [];
        } catch { return []; }
    }

    getSemanticDiagnostics(): ts.Diagnostic[] {
        try {
            return this._service.getSemanticDiagnostics(VIRTUAL_FILENAME) ?? [];
        } catch { return []; }
    }

    getEncodedSemanticClassifications(
        start: number, length: number
    ): ts.Classifications {
        try {
            return this._service.getEncodedSemanticClassifications(
                VIRTUAL_FILENAME, { start, length },
                ts.SemanticClassificationFormat.TwentyTwenty
            );
        } catch {
            return { spans: [], endOfLineState: ts.EndOfLineState.None };
        }
    }

    dispose(): void {
        this._service.dispose();
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Singleton
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
// Zone helpers
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

// ─────────────────────────────────────────────────────────────────────────────
// ts.DiagnosticCategory → vscode.DiagnosticSeverity
// ─────────────────────────────────────────────────────────────────────────────
export function tsSeverityToVs(category: ts.DiagnosticCategory): vscode.DiagnosticSeverity {
    switch (category) {
        case ts.DiagnosticCategory.Error:       return vscode.DiagnosticSeverity.Error;
        case ts.DiagnosticCategory.Warning:     return vscode.DiagnosticSeverity.Warning;
        case ts.DiagnosticCategory.Suggestion:  return vscode.DiagnosticSeverity.Hint;
        case ts.DiagnosticCategory.Message:     return vscode.DiagnosticSeverity.Information;
        default:                                return vscode.DiagnosticSeverity.Warning;
    }
}