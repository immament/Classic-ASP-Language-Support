import * as vscode from 'vscode';

export interface AspFormatterSettings {
    keywordCase: string;
    useTabs: boolean;
    indentSize: number;
    aspTagsOnSameLine: boolean;
}

export function getAspSettings(): AspFormatterSettings {
    const config = vscode.workspace.getConfiguration('aspLanguageSupport');
    return {
        keywordCase:       config.get<string>('keywordCase', 'PascalCase'),
        useTabs:           config.get<boolean>('useTabs', false),
        indentSize:        config.get<number>('indentSize', 2),
        aspTagsOnSameLine: config.get<boolean>('aspTagsOnSameLine', false),
    };
}

// ─── Static lookup tables (built once at module load) ──────────────────────

const KEYWORDS_WITH_PROPER_CASING: Record<string, string> = {
    'elseif': 'ElseIf', 'redim': 'ReDim', 'byval': 'ByVal', 'byref': 'ByRef',
    'isnull': 'IsNull', 'isempty': 'IsEmpty', 'isnumeric': 'IsNumeric',
    'isarray': 'IsArray', 'isobject': 'IsObject', 'isdate': 'IsDate',
    'readonly': 'ReadOnly', 'writeonly': 'WriteOnly', 'typename': 'TypeName',
    'vartype': 'VarType', 'getobject': 'GetObject', 'createobject': 'CreateObject',
    'getref': 'GetRef', 'endif': 'EndIf', 'endsub': 'EndSub',
    'endfunction': 'EndFunction', 'endwith': 'EndWith', 'endselect': 'EndSelect',
    'endclass': 'EndClass', 'endproperty': 'EndProperty', 'exitfor': 'ExitFor',
    'exitdo': 'ExitDo', 'exitsub': 'ExitSub', 'exitfunction': 'ExitFunction',
    'exitproperty': 'ExitProperty', 'onerror': 'OnError',
    'querystring': 'QueryString', 'servervariables': 'ServerVariables',
    'totalbytes': 'TotalBytes', 'binaryread': 'BinaryRead',
    'clientcertificate': 'ClientCertificate', 'contenttype': 'ContentType',
    'addheader': 'AddHeader', 'appendtolog': 'AppendToLog',
    'binarywrite': 'BinaryWrite', 'cacheecontrol': 'CacheControl',
    'charset': 'Charset', 'clearheaders': 'ClearHeaders',
    'contentlength': 'ContentLength', 'expires': 'Expires',
    'expiresabsolute': 'ExpiresAbsolute', 'isclientconnected': 'IsClientConnected',
    'pics': 'PICS', 'status': 'Status', 'mappath': 'MapPath',
    'scripttimeout': 'ScriptTimeout', 'htmlencode': 'HTMLEncode',
    'urlencode': 'URLEncode', 'createtextfile': 'CreateTextFile',
    'opentextfile': 'OpenTextFile', 'getlasterror': 'GetLastError',
    'sessionid': 'SessionID', 'timeout': 'Timeout', 'codepage': 'CodePage',
    'lcid': 'LCID', 'filesystemobject': 'FileSystemObject',
    'getfile': 'GetFile', 'getfolder': 'GetFolder', 'getdrive': 'GetDrive',
    'fileexists': 'FileExists', 'folderexists': 'FolderExists',
    'driveexists': 'DriveExists', 'getfilename': 'GetFileName',
    'getbasename': 'GetBaseName', 'getextensionname': 'GetExtensionName',
    'getparentfoldername': 'GetParentFolderName', 'getdrivename': 'GetDriveName',
    'getabsolutepathname': 'GetAbsolutePathName', 'buildpath': 'BuildPath',
    'getspecialfolder': 'GetSpecialFolder', 'gettempname': 'GetTempName',
    'deletefile': 'DeleteFile', 'deletefolder': 'DeleteFolder',
    'movefile': 'MoveFile', 'movefolder': 'MoveFolder',
    'copyfile': 'CopyFile', 'copyfolder': 'CopyFolder', 'createfolder': 'CreateFolder',
    'writeline': 'WriteLine', 'writeblanklines': 'WriteBlankLines',
    'readline': 'ReadLine', 'readall': 'ReadAll', 'atendofstream': 'AtEndOfStream',
    'atendofline': 'AtEndOfLine', 'skipline': 'SkipLine', 'closetext': 'CloseText',
    'datelastmodified': 'DateLastModified', 'datelastaccessed': 'DateLastAccessed',
    'datecreated': 'DateCreated', 'parentfolder': 'ParentFolder',
    'shortname': 'ShortName', 'shortpath': 'ShortPath', 'rootfolder': 'RootFolder',
    'recordset': 'Recordset', 'movenext': 'MoveNext', 'movefirst': 'MoveFirst',
    'movelast': 'MoveLast', 'moveprevious': 'MovePrevious', 'addnew': 'AddNew',
    'recordcount': 'RecordCount', 'pagesize': 'PageSize', 'pagecount': 'PageCount',
    'absolutepage': 'AbsolutePage', 'absoluteposition': 'AbsolutePosition',
    'cursortype': 'CursorType', 'cursorlocation': 'CursorLocation',
    'locktype': 'LockType', 'commandtext': 'CommandText', 'commandtype': 'CommandType',
    'connectionstring': 'ConnectionString', 'begintrans': 'BeginTrans',
    'committrans': 'CommitTrans', 'rollbacktrans': 'RollbackTrans',
};

const VBSCRIPT_FUNCTIONS: Record<string, string> = {
    'cbool': 'CBool', 'cbyte': 'CByte', 'ccur': 'CCur', 'cdate': 'CDate',
    'cdbl': 'CDbl', 'cint': 'CInt', 'clng': 'CLng', 'csng': 'CSng',
    'cstr': 'CStr', 'cvar': 'CVar',
    'isarray': 'IsArray', 'isdate': 'IsDate', 'isempty': 'IsEmpty',
    'isnull': 'IsNull', 'isnumeric': 'IsNumeric', 'isobject': 'IsObject',
    'lcase': 'LCase', 'ucase': 'UCase', 'ltrim': 'LTrim', 'rtrim': 'RTrim',
    'instr': 'InStr', 'instrrev': 'InStrRev', 'strreverse': 'StrReverse',
    'strcomp': 'StrComp',
    'dateserial': 'DateSerial', 'timeserial': 'TimeSerial',
    'datevalue': 'DateValue', 'timevalue': 'TimeValue',
    'dateadd': 'DateAdd', 'datediff': 'DateDiff', 'datepart': 'DatePart',
    'formatdatetime': 'FormatDateTime', 'formatnumber': 'FormatNumber',
    'formatcurrency': 'FormatCurrency', 'formatpercent': 'FormatPercent',
    'monthname': 'MonthName', 'weekdayname': 'WeekdayName',
    'lbound': 'LBound', 'ubound': 'UBound',
    'createobject': 'CreateObject', 'getobject': 'GetObject',
    'msgbox': 'MsgBox', 'inputbox': 'InputBox',
    'typename': 'TypeName', 'vartype': 'VarType', 'getref': 'GetRef',
    'eval': 'Eval', 'loadpicture': 'LoadPicture', 'scriptengine': 'ScriptEngine',
    'scriptenginebuildversion': 'ScriptEngineBuildVersion',
    'scriptenginemajorversion': 'ScriptEngineMajorVersion',
    'scriptengineminorversion': 'ScriptEngineMinorVersion',
    'rgb': 'RGB', 'escape': 'Escape', 'unescape': 'Unescape',
    'getlocale': 'GetLocale', 'setlocale': 'SetLocale',
};

const KEYWORDS_SORTED: string[] = [
    'if', 'then', 'else', 'elseif', 'end if', 'select case', 'case', 'case else',
    'end select', 'for', 'to', 'step', 'next', 'for each', 'in', 'while', 'wend',
    'do', 'loop', 'until', 'exit do', 'exit for', 'sub', 'end sub', 'function',
    'end function', 'call', 'exit sub', 'exit function', 'dim', 'redim', 'preserve',
    'const', 'private', 'public', 'static', 'class', 'end class', 'new', 'set',
    'property get', 'property let', 'property set', 'end property',
    'on error resume next', 'on error goto 0', 'err', 'error',
    'and', 'or', 'not', 'xor', 'eqv', 'imp', 'is', 'like',
    'nothing', 'null', 'empty', 'true', 'false',
    'option explicit', 'randomize', 'with', 'end with', 'exit', 'mod',
    'byval', 'byref', 'default', 'erase', 'let', 'resume', 'stop', 'get', 'put',
    'open', 'close', 'input', 'output', 'append', 'binary', 'random', 'as',
    'len', 'mid', 'left', 'right', 'trim', 'replace', 'split', 'join', 'filter',
    'string', 'space', 'chr', 'asc', 'int', 'fix', 'abs', 'sgn', 'sqr', 'exp',
    'log', 'sin', 'cos', 'tan', 'atn', 'round', 'rnd',
    'array', 'date', 'time', 'now', 'timer',
    'year', 'month', 'day', 'weekday', 'hour', 'minute', 'second',
    'response', 'request', 'server', 'session', 'application',
    'write', 'redirect', 'querystring', 'form', 'servervariables',
    'cookies', 'mappath', 'createtextfile', 'opentextfile', 'writeline',
    'readline', 'readall', 'atendofstream', 'filesystemobject', 'scripting',
    'dictionary', 'add', 'exists', 'items', 'keys', 'remove', 'removeall',
    'count', 'item', 'key',
].sort((a, b) => b.length - a.length);

const KEYWORD_REGEXES: RegExp[] = KEYWORDS_SORTED.map(
    kw => new RegExp('\\b' + kw.replace(/\s+/g, '\\s+') + '\\b', 'gi')
);

const PROPER_CASING_REGEXES: Array<{ re: RegExp; replacement: string }> =
    Object.entries(KEYWORDS_WITH_PROPER_CASING).map(([lower, proper]) => ({
        re: new RegExp('\\b' + lower + '\\b', 'gi'),
        replacement: proper,
    }));

const VBSCRIPT_FUNCTION_REGEXES: Array<{ re: RegExp; replacement: string }> =
    Object.entries(VBSCRIPT_FUNCTIONS).map(([lower, proper]) => ({
        re: new RegExp('\\b' + lower + '\\b', 'gi'),
        replacement: proper,
    }));

const HANDLED_KEYWORDS = new Set<string>([
    ...Object.keys(VBSCRIPT_FUNCTIONS),
    ...Object.keys(KEYWORDS_WITH_PROPER_CASING),
]);

// ─── Public API ────────────────────────────────────────────────────────────

export interface FormatBlockResult {
    formatted: string;
    endLevel: number;  // VBScript indent level at end of block, for threading to next block
}

// Formats a single ASP block and returns both the formatted string and the
// ending indent level so callers can thread state across multiple blocks.
export function formatSingleAspBlock(
    block: string,
    settings: AspFormatterSettings,
    htmlIndent: string = '',
    startLevel: number = 0,
): FormatBlockResult {
    const trimmedBlock = block.trim();

    // <%= expression %> — no indent tracking needed
    if (trimmedBlock.startsWith('<%=') || trimmedBlock.startsWith('<% =')) {
        const content = trimmedBlock.startsWith('<%=')
            ? trimmedBlock.substring(3, trimmedBlock.length - 2).trim()
            : trimmedBlock.substring(4, trimmedBlock.length - 2).trim();
        return {
            formatted: htmlIndent + '<%= ' + applyKeywordCase(content, settings.keywordCase) + ' %>',
            endLevel: startLevel,
        };
    }

    // Single-line block — content sits at the same visual level as <% and %>
    if (!block.includes('\n')) {
        const content = block.substring(2, block.length - 2).trim();
        const formattedContent = applyKeywordCase(content, settings.keywordCase);
        const selectCaseStack: number[] = [];
        const afterBefore = applyIndentBefore(content, startLevel, selectCaseStack).level;
        const endLevel = applyIndentAfter(content, afterBefore, selectCaseStack);

        if (settings.aspTagsOnSameLine) {
            return { formatted: htmlIndent + '<% ' + formattedContent + ' %>', endLevel };
        } else {
            return {
                formatted: htmlIndent + '<%\n' + htmlIndent + formattedContent + '\n' + htmlIndent + '%>',
                endLevel,
            };
        }
    }

    return formatMultiLineAspBlock(block, settings, htmlIndent, startLevel);
}

// ─── Private helpers ───────────────────────────────────────────────────────

function formatMultiLineAspBlock(
    block: string,
    settings: AspFormatterSettings,
    htmlIndent: string,
    startLevel: number,
): FormatBlockResult {
    const lines = block.split('\n');
    const formattedLines: string[] = [];
    let aspIndentLevel = startLevel;
    let previousLineHadContinuation = false;
    let continuationAlignColumn = 0;
    let inMultilineString = false;
    let sqlBaseIndent = 0;
    let isInSQLBlock = false;
    // Stack tracking base indent level at each Select Case opener, used to
    // correctly indent Case labels and their bodies.
    const selectCaseStack: number[] = [];

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const trimmedLine = line.trim();

        // ── Opening tag line (%< ...) ────────────────────────────────────
        if (trimmedLine.startsWith('<%')) {
            if (trimmedLine === '<%') {
                formattedLines.push(htmlIndent + '<%');
                previousLineHadContinuation = false;
                isInSQLBlock = false;
                continue;
            }

            const content = trimmedLine.substring(2).trim();
            if (content) {
                const indentResult = applyIndentBefore(content, aspIndentLevel, selectCaseStack);
                aspIndentLevel = indentResult.level;

                const aspIndent     = getIndentString(aspIndentLevel, settings.useTabs, settings.indentSize);
                const formattedContent = applyKeywordCase(content, settings.keywordCase);

                if (settings.aspTagsOnSameLine) {
                    formattedLines.push(htmlIndent + '<% ' + formattedContent);
                } else {
                    formattedLines.push(htmlIndent + '<%');
                    formattedLines.push(htmlIndent + aspIndent + formattedContent);
                }

                handleContinuation(formattedContent, htmlIndent, aspIndent, {
                    get: () => ({ previousLineHadContinuation, continuationAlignColumn, inMultilineString, isInSQLBlock, sqlBaseIndent }),
                    set: (v) => { previousLineHadContinuation = v.previousLineHadContinuation; continuationAlignColumn = v.continuationAlignColumn; inMultilineString = v.inMultilineString; isInSQLBlock = v.isInSQLBlock; sqlBaseIndent = v.sqlBaseIndent; },
                });

                aspIndentLevel = applyIndentAfter(content, aspIndentLevel, selectCaseStack);
            }
            continue;
        }

        // ── Closing tag line (%>) ────────────────────────────────────────
        if (trimmedLine === '%>' || trimmedLine.endsWith('%>')) {
            if (trimmedLine === '%>') {
                formattedLines.push(htmlIndent + '%>');
                previousLineHadContinuation = false;
                inMultilineString = false;
                isInSQLBlock = false;
                continue;
            }

            const content = trimmedLine.substring(0, trimmedLine.length - 2).trim();
            if (content) {
                const indentResult = applyIndentBefore(content, aspIndentLevel, selectCaseStack);
                aspIndentLevel = indentResult.level;

                const aspIndent     = getIndentString(aspIndentLevel, settings.useTabs, settings.indentSize);
                const formattedContent = applyKeywordCase(content, settings.keywordCase);

                if (settings.aspTagsOnSameLine) {
                    formattedLines.push(htmlIndent + aspIndent + formattedContent + ' %>');
                } else {
                    formattedLines.push(htmlIndent + aspIndent + formattedContent);
                    formattedLines.push(htmlIndent + '%>');
                }

                aspIndentLevel = applyIndentAfter(content, aspIndentLevel, selectCaseStack);
            } else {
                formattedLines.push(htmlIndent + '%>');
            }
            previousLineHadContinuation = false;
            inMultilineString = false;
            isInSQLBlock = false;
            continue;
        }

        // ── Empty line ───────────────────────────────────────────────────
        if (!trimmedLine) {
            formattedLines.push('');
            if (!inMultilineString) {
                previousLineHadContinuation = false;
                isInSQLBlock = false;
            }
            continue;
        }

        // ── Comment line ─────────────────────────────────────────────────
        if (trimmedLine.startsWith("'")) {
            // Align comment with the next non-comment, non-empty line's indent.
            let commentLevel = aspIndentLevel;
            for (let j = i + 1; j < lines.length; j++) {
                const next = lines[j].trim();
                if (next && !next.startsWith("'") && next !== '%>') {
                    const r = applyIndentBefore(next, aspIndentLevel, [...selectCaseStack]);
                    commentLevel = r.level;
                    break;
                }
            }
            formattedLines.push(htmlIndent + getIndentString(commentLevel, settings.useTabs, settings.indentSize) + trimmedLine);
            previousLineHadContinuation = false;
            isInSQLBlock = false;
            continue;
        }

        // ── String continuation line ─────────────────────────────────────
        if (previousLineHadContinuation && trimmedLine.startsWith('"')) {
            // Only use SQL mode if the FIRST line of the continuation set was
            // detected as SQL (isInSQLBlock). Do NOT re-test the continuation
            // line itself, since its SQL keywords are inside a string literal.
            if (isInSQLBlock) {
                const originalIndentSize = line.length - line.trimStart().length;
                const relativeIndent = originalIndentSize - sqlBaseIndent;
                const extraLevel = relativeIndent > 0 ? 1 : 0;
                const aspIndent = getIndentString(aspIndentLevel + 1 + extraLevel, settings.useTabs, settings.indentSize);
                formattedLines.push(htmlIndent + aspIndent + trimmedLine);
            } else if (continuationAlignColumn === -1) {
                const aspIndent = getIndentString(aspIndentLevel + 1, settings.useTabs, settings.indentSize);
                formattedLines.push(htmlIndent + aspIndent + trimmedLine);
            } else {
                formattedLines.push(' '.repeat(continuationAlignColumn) + trimmedLine);
            }

            if (!trimmedLine.trimEnd().endsWith('_')) {
                previousLineHadContinuation = false;
                inMultilineString = false;
                isInSQLBlock = false;
            }
            continue;
        }

        // ── Normal VBScript line ─────────────────────────────────────────
        const indentResult = applyIndentBefore(trimmedLine, aspIndentLevel, selectCaseStack);
        aspIndentLevel = indentResult.level;

        const aspIndent      = getIndentString(aspIndentLevel, settings.useTabs, settings.indentSize);
        const formattedContent = applyKeywordCase(trimmedLine, settings.keywordCase);

        handleContinuation(formattedContent, htmlIndent, aspIndent, {
            get: () => ({ previousLineHadContinuation, continuationAlignColumn, inMultilineString, isInSQLBlock, sqlBaseIndent }),
            set: (v) => { previousLineHadContinuation = v.previousLineHadContinuation; continuationAlignColumn = v.continuationAlignColumn; inMultilineString = v.inMultilineString; isInSQLBlock = v.isInSQLBlock; sqlBaseIndent = v.sqlBaseIndent; },
        });

        formattedLines.push(htmlIndent + aspIndent + formattedContent);
        aspIndentLevel = applyIndentAfter(trimmedLine, aspIndentLevel, selectCaseStack);
    }

    return { formatted: formattedLines.join('\n'), endLevel: aspIndentLevel };
}

// ─── Indent helpers ────────────────────────────────────────────────────────

// Applies indent-decreasing rules BEFORE printing the line.
// Returns the new level to print at.
function applyIndentBefore(
    line: string,
    level: number,
    selectCaseStack: number[],
): { level: number } {
    const lowerLine = removeStringsFromLine(line).toLowerCase().trim();

    // End Select: pop the Select Case stack and drop two levels.
    if (/^\s*end\s+select\b/.test(lowerLine)) {
        if (selectCaseStack.length > 0) {
            return { level: selectCaseStack.pop()! };
        }
        return { level: Math.max(0, level - 2) };
    }

    // Case / Case Else: jump back to the Case label level (baseLevel + 1).
    if (/^\s*case(\s|$)/.test(lowerLine)) {
        if (selectCaseStack.length > 0) {
            return { level: selectCaseStack[selectCaseStack.length - 1] + 1 };
        }
        return { level: Math.max(0, level - 1) };
    }

    // Standard dedent-before keywords
    if (
        /^\s*end\s+(if|sub|function|with|class|property)\b/.test(lowerLine) ||
        /^\s*loop(\s|$)/.test(lowerLine) ||
        /^\s*next(\s|$)/.test(lowerLine) ||
        /^\s*wend(\s|$)/.test(lowerLine) ||
        /^\s*else(\s|$)/.test(lowerLine) ||
        /^\s*elseif\b/.test(lowerLine)
    ) {
        return { level: Math.max(0, level - 1) };
    }

    return { level };
}

// Applies indent-increasing rules AFTER printing the line.
function applyIndentAfter(
    line: string,
    level: number,
    selectCaseStack: number[],
): number {
    const lowerLine = removeStringsFromLine(line).toLowerCase().trim();

    // Single-line If ... Then <statement> — no indent change
    if (/\bif\b.*\bthen\b\s+\S/.test(lowerLine)) return level;

    // Select Case: push current level onto stack, jump to level+1 for Case labels.
    if (/\bselect\s+case\b/.test(lowerLine)) {
        selectCaseStack.push(level);
        return level + 1; // Case labels will be at level+1, bodies at level+2
    }

    // Case / Case Else: body is one deeper than the Case label.
    if (/^\s*case(\s|$)/.test(lowerLine)) {
        return level + 1;
    }

    // Standard indent-after keywords
    if (
        /\bif\b.*\bthen\b/.test(lowerLine) ||
        /\bfor\b\s+\w+\s*=/.test(lowerLine) ||
        /\bfor\s+each\b/.test(lowerLine) ||
        /\bwhile\b/.test(lowerLine) ||
        /\bdo\b(\s+while|\s+until)?(\s|$)/.test(lowerLine) ||
        /\bsub\b\s+\w+/.test(lowerLine) ||
        /\bfunction\b\s+\w+/.test(lowerLine) ||
        /\bwith\b/.test(lowerLine) ||
        /\bclass\b\s+\w+/.test(lowerLine) ||
        /\bproperty\s+(get|let|set)\b/.test(lowerLine) ||
        /^\s*else(\s|$)/.test(lowerLine) ||
        /^\s*elseif\b.*\bthen\b/.test(lowerLine)
    ) {
        return level + 1;
    }

    return level;
}

// Handles line-continuation state updates in one place.
interface ContinuationState {
    previousLineHadContinuation: boolean;
    continuationAlignColumn: number;
    inMultilineString: boolean;
    isInSQLBlock: boolean;
    sqlBaseIndent: number;
}

function handleContinuation(
    formattedContent: string,
    htmlIndent: string,
    aspIndent: string,
    state: { get: () => ContinuationState; set: (v: ContinuationState) => void },
): void {
    const hasContinuation = formattedContent.trimEnd().endsWith('_');
    const s = state.get();

    if (hasContinuation) {
        const col = calculateContinuationColumn(formattedContent, htmlIndent, aspIndent);
        const isSql = isSQLStatement(formattedContent);
        state.set({
            previousLineHadContinuation: true,
            continuationAlignColumn: col,
            inMultilineString: true,
            isInSQLBlock: isSql,
            sqlBaseIndent: isSql ? (htmlIndent + aspIndent).length : s.sqlBaseIndent,
        });
    } else {
        state.set({
            ...s,
            previousLineHadContinuation: false,
            inMultilineString: false,
            isInSQLBlock: false,
        });
    }
}

// Returns column to align continuation strings, or -1 if no string on the line.
function calculateContinuationColumn(line: string, htmlIndent: string, aspIndent: string): number {
    const fullIndentLen = (htmlIndent + aspIndent).length;
    const trimmed = line.trim();

    const equalsPos = trimmed.indexOf('=');
    if (equalsPos !== -1) {
        const afterEquals = trimmed.substring(equalsPos + 1).trim();
        if (afterEquals.startsWith('"')) {
            const quoteOffset = trimmed.substring(equalsPos).indexOf('"');
            return fullIndentLen + equalsPos + quoteOffset;
        }
    }

    const quotePos = trimmed.indexOf('"');
    if (quotePos !== -1) return fullIndentLen + quotePos;

    return -1; // No string — caller should use +1 indent level
}

function getIndentString(level: number, useTabs: boolean, indentSize: number): string {
    return useTabs ? '\t'.repeat(Math.max(0, level)) : ' '.repeat(Math.max(0, level) * indentSize);
}

// ─── Keyword/operator formatting ───────────────────────────────────────────

function applyKeywordCase(code: string, caseStyle: string): string {
    return splitByStrings(code).map(part => {
        if (part.isString) return part.text;
        let s = formatKeywordsInText(part.text, caseStyle);
        s = formatOperators(s);
        s = formatCommas(s);
        return s;
    }).join('');
}

function formatKeywordsInText(text: string, caseStyle: string): string {
    let result = text;

    if (caseStyle === 'PascalCase') {
        for (const { re, replacement } of PROPER_CASING_REGEXES) {
            result = result.replace(re, replacement);
        }
    }

    for (const { re, replacement } of VBSCRIPT_FUNCTION_REGEXES) {
        result = result.replace(re, replacement);
    }

    for (let idx = 0; idx < KEYWORDS_SORTED.length; idx++) {
        if (HANDLED_KEYWORDS.has(KEYWORDS_SORTED[idx].toLowerCase())) continue;
        result = result.replace(KEYWORD_REGEXES[idx], m => formatKeyword(m, caseStyle));
    }

    return result;
}

function formatKeyword(keyword: string, caseStyle: string): string {
    switch (caseStyle) {
        case 'lowercase': return keyword.toLowerCase();
        case 'UPPERCASE': return keyword.toUpperCase();
        default:          return keyword.split(' ')
            .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
            .join(' ');
    }
}

function formatOperators(code: string): string {
    return splitByStrings(code).map(part =>
        part.isString ? part.text : formatOperatorsInText(part.text)
    ).join('');
}

function formatOperatorsInText(text: string): string {
    let r = text;
    // Compound operators first to avoid them being split by single-char rules.
    r = r.replace(/\s*<>\s*/g, ' <> ');
    r = r.replace(/\s*<=\s*/g, ' <= ');
    r = r.replace(/\s*>=\s*/g, ' >= ');
    r = r.replace(/\s*=\s*/g,  ' = ');
    r = r.replace(/\s*\+\s*/g, ' + ');
    r = r.replace(/\s*-\s*/g,  ' - ');
    r = r.replace(/\s*\*\s*/g, ' * ');
    r = r.replace(/\s*\/\s*/g, ' / ');
    r = r.replace(/\s*&\s*/g,  ' & ');
    // < and > last, skip those already part of a compound operator.
    r = r.replace(/(?<![<>!])\s*<\s*(?![>=])/g, ' < ');
    r = r.replace(/(?<![<>])\s*>\s*(?![=])/g,   ' > ');
    return r;
}

function formatCommas(code: string): string {
    return splitByStrings(code).map(part =>
        part.isString ? part.text : part.text.replace(/,(?!\s)/g, ', ')
    ).join('');
}

function isSQLStatement(line: string): boolean {
    // Check only outside strings for SQL keywords.
    const stripped = removeStringsFromLine(line);
    return /\b(SELECT|INSERT|UPDATE|DELETE|FROM|WHERE|JOIN|ORDER\s+BY|GROUP\s+BY|UNION|CREATE|DROP|ALTER|INNER|LEFT|RIGHT|OUTER|HAVING|DISTINCT|VALUES|INTO)\b/i.test(stripped);
}

function removeStringsFromLine(line: string): string {
    let result = '';
    let inString = false;

    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
            if (i + 1 < line.length && line[i + 1] === '"') { i++; continue; }
            inString = !inString;
        } else if (!inString) {
            result += char;
        }
    }

    return result;
}

function splitByStrings(code: string): Array<{ text: string; isString: boolean }> {
    const parts: Array<{ text: string; isString: boolean }> = [];
    let current = '';
    let inString = false;

    for (let i = 0; i < code.length; i++) {
        const char = code[i];

        if (char === '"') {
            if (i + 1 < code.length && code[i + 1] === '"') {
                current += '""';
                i++;
                continue;
            }
            if (inString) {
                current += char;
                parts.push({ text: current, isString: true });
                current = '';
                inString = false;
            } else {
                if (current) parts.push({ text: current, isString: false });
                current = char;
                inString = true;
            }
        } else {
            current += char;
        }
    }

    if (current) parts.push({ text: current, isString: inString });
    return parts;
}