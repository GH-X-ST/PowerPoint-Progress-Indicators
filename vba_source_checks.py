"""Small source guards for VBA declarations, not a VBA parser or compiler.

The reserved-name check follows MS-VBAL sections 3.3.5.2 and 5.2.3.1.1.
It inspects ordinary variable, constant, and parameter declarations. It does
not resolve Office types, validate procedure bodies, or measure compiled code.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterator


# MS-VBAL 3.3.5.2. Built-in function names outside this list, such as Left,
# are not reserved identifiers and must not be rejected by this check.
RESERVED_IDENTIFIERS = frozenset("""
Call Case Close Const Declare DefBool DefByte DefCur DefDate DefDbl DefInt
DefLng DefLngLng DefLngPtr DefObj DefSng DefStr DefVar Dim Do Else ElseIf End
EndIf Enum Erase Event Exit For Friend Function Get Global GoSub GoTo If
Implements Input Let Lock Loop LSet Next On Open Option Print Private Public
Put RaiseEvent ReDim Resume Return RSet Seek Select Set Static Stop Sub Type
Unlock Wend While With Write Rem Any As ByRef ByVal Each In New Shared Until
WithEvents Optional ParamArray Preserve Spc Tab Then To AddressOf And Eqv Imp
Is Like Mod Not Or TypeOf Xor Abs CBool CByte CCur CDate CDbl CDec CInt CLng
CLngLng CLngPtr CSng CStr CVar CVErr Date Debug DoEvents Fix Int Len LenB Me
PSet Scale Sgn String Array Circle InputB LBound UBound Boolean Byte Currency
Double Integer Long LongLong LongPtr Single Variant True False Nothing Empty
Null Attribute LINEINPUT VB_Base VB_Control VB_Creatable VB_Customizable
VB_Description VB_Exposed VB_Ext_KEY VB_GlobalNameSpace VB_HelpID
VB_Invoke_Func VB_Invoke_Property VB_Invoke_PropertyPut VB_Invoke_PropertyPutRef
VB_MemberFlags VB_Name VB_PredeclaredId VB_ProcData VB_TemplateDerived
VB_UserMemId VB_VarDescription VB_VarHelpID VB_VarMemberFlags VB_VarProcData
VB_VarUserMemId CDecl Decimal DefDec
""".lower().split())

_TYPE_CHARACTERS = "%&^!#$@"
_ACCESS = {"public", "private", "friend", "global"}


@dataclass(frozen=True)
class DeclarationIssue:
    filename: str
    line: int
    identifier: str
    context: str


class VBASourceCheckError(ValueError):
    """An ordinary VBA declaration uses a reserved identifier."""

    def __init__(self, issues: list[DeclarationIssue]):
        self.issues = issues
        detail = "\n".join(
            f"{item.filename}:{item.line}: reserved identifier "
            f"{item.identifier!r} in {item.context} declaration"
            for item in issues
        )
        super().__init__(detail)


@dataclass(frozen=True)
class _Token:
    text: str
    line: int
    kind: str = "symbol"

    @property
    def lower(self) -> str:
        return self.text.lower()


def _statements(source: str) -> Iterator[list[_Token]]:
    """Tokenize declarations while retaining source line numbers.

    Quoted strings, date literals, bracketed names, comments, continuations,
    and statement separators are handled before looking for declaration names.
    """
    source = source.replace("\r\n", "\n").replace("\r", "\n")
    tokens: list[_Token] = []
    index = 0
    line = 1
    while index < len(source):
        character = source[index]
        if character == "\n":
            if tokens and tokens[-1].text == "_":
                tokens.pop()
            else:
                if tokens:
                    yield tokens
                tokens = []
            line += 1
            index += 1
        elif character.isspace() or character == "\ufeff":
            index += 1
        elif character == "'":
            newline = source.find("\n", index)
            index = len(source) if newline < 0 else newline
        elif character == '"':
            token_line = line
            index += 1
            while index < len(source):
                if source[index] == '"':
                    index += 1
                    if index < len(source) and source[index] == '"':
                        index += 1
                        continue
                    break
                # A malformed multiline literal is left for the VBA compiler.
                line += source[index] == "\n"
                index += 1
            tokens.append(_Token('""', token_line, "literal"))
        elif character == "[":
            end = source.find("]", index + 1)
            newline = source.find("\n", index + 1)
            if end >= 0 and (newline < 0 or end < newline):
                tokens.append(_Token(source[index:end + 1], line, "bracketed"))
                index = end + 1
            else:
                tokens.append(_Token(character, line))
                index += 1
        elif character == "#" and tokens:
            end = source.find("#", index + 1)
            newline = source.find("\n", index + 1)
            if end >= 0 and (newline < 0 or end < newline):
                tokens.append(_Token("##", line, "literal"))
                index = end + 1
            else:
                tokens.append(_Token(character, line))
                index += 1
        elif character.isalpha() or character == "_":
            end = index + 1
            while end < len(source) and (source[end].isalnum() or source[end] == "_"):
                end += 1
            if end < len(source) and source[end] in _TYPE_CHARACTERS:
                end += 1
            token = _Token(source[index:end], line, "identifier")
            at_statement_start = not tokens or (len(tokens) == 1 and tokens[0].kind == "number")
            if token.lower == "rem" and at_statement_start:
                newline = source.find("\n", end)
                index = len(source) if newline < 0 else newline
            else:
                tokens.append(token)
                index = end
        elif character.isdigit():
            end = index + 1
            while end < len(source) and (source[end].isalnum() or source[end] in "."):
                end += 1
            tokens.append(_Token(source[index:end], line, "number"))
            index = end
        elif character == ":" and source[index:index + 2] != ":=":
            if tokens:
                yield tokens
            tokens = []
            index += 1
        elif source[index:index + 2] == ":=":
            tokens.append(_Token(":=", line))
            index += 2
        else:
            tokens.append(_Token(character, line))
            index += 1
    if tokens:
        yield tokens


def _split_top_level(tokens: list[_Token], delimiter: str = ",") -> Iterator[list[_Token]]:
    start = 0
    depth = 0
    for index, token in enumerate(tokens):
        if token.text == "(":
            depth += 1
        elif token.text == ")":
            depth -= 1
        elif token.text == delimiter and depth == 0:
            yield tokens[start:index]
            start = index + 1
    yield tokens[start:]


def _parameter_tokens(tokens: list[_Token]) -> list[_Token]:
    """Return the first parenthesized parameter list of a procedure header."""
    start = next((i for i, token in enumerate(tokens) if token.text == "("), -1)
    if start < 0:
        return []
    depth = 1
    for end in range(start + 1, len(tokens)):
        if tokens[end].text == "(":
            depth += 1
        elif tokens[end].text == ")":
            depth -= 1
            if depth == 0:
                return tokens[start + 1:end]
    return []


def _declarations(tokens: list[_Token]) -> Iterator[tuple[_Token, str]]:
    if tokens and tokens[0].kind == "number":
        tokens = tokens[1:]
    if not tokens:
        return

    # Declarations can occur inside a single-line If statement. Only split
    # its outer Then/Else branches, not tokens in a procedure parameter list.
    if tokens[0].lower in {"if", "elseif", "else"}:
        depth = 0
        branches: list[int] = []
        for index, token in enumerate(tokens):
            depth += token.text == "("
            depth -= token.text == ")"
            if depth == 0 and token.lower in {"then", "else"}:
                branches.append(index)
        for start, end in zip(branches, branches[1:] + [len(tokens)]):
            yield from _declarations(tokens[start + 1:end])
        return

    index = 0
    if tokens[index].lower in _ACCESS:
        index += 1
    if index >= len(tokens):
        return
    first = tokens[index].lower
    if first == "static" and index + 1 < len(tokens) and tokens[index + 1].lower in {"sub", "function", "property"}:
        index += 1
        first = tokens[index].lower
    if first == "declare":
        index += 1
        if index < len(tokens) and tokens[index].lower == "ptrsafe":
            index += 1
        if index >= len(tokens):
            return
        first = tokens[index].lower
    if first in {"sub", "function", "property", "event"}:
        for parameter in _split_top_level(_parameter_tokens(tokens[index:])):
            position = 0
            if parameter and parameter[position].lower == "optional":
                position += 1
            if position < len(parameter) and parameter[position].lower in {"byval", "byref"}:
                position += 1
            if position < len(parameter) and parameter[position].lower == "paramarray":
                position += 1
            if position < len(parameter):
                yield parameter[position], "parameter"
        return
    if first in {"type", "enum"}:
        return

    if first == "const":
        index += 1
        context = "constant"
    elif first in {"dim", "static"}:
        index += 1
        context = "variable"
    elif index > 0:
        context = "variable"
    else:
        return
    if index < len(tokens) and tokens[index].lower == "withevents":
        index += 1
    for declaration in _split_top_level(tokens[index:]):
        if declaration:
            yield declaration[0], context


def find_reserved_declarations(source: str, filename: str = "<source>") -> list[DeclarationIssue]:
    """Find reserved ordinary declaration names, ignoring literals and members."""
    issues: list[DeclarationIssue] = []
    for statement in _statements(source):
        for token, context in _declarations(statement):
            if token.kind != "identifier":
                continue
            identifier = token.text.rstrip(_TYPE_CHARACTERS)
            if identifier.lower() in RESERVED_IDENTIFIERS:
                issues.append(DeclarationIssue(filename, token.line, token.text, context))
    return issues


def validate_vba_sources(source_dir: Path) -> dict:
    """Check .bas/.cls/.frm sources, raising VBASourceCheckError on findings."""
    source_dir = Path(source_dir)
    paths = sorted(path for path in source_dir.iterdir() if path.suffix.lower() in {".bas", ".cls", ".frm"})
    if not paths:
        raise ValueError(f"No VBA source modules found in {source_dir}")
    issues: list[DeclarationIssue] = []
    for path in paths:
        issues.extend(find_reserved_declarations(path.read_text(encoding="utf-8-sig"), path.name))
    if issues:
        raise VBASourceCheckError(issues)
    return {"modules_checked": len(paths), "reserved_declaration_issues": [asdict(issue) for issue in issues],
            "native_vba_compilation_tested": False}
