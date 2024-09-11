Attribute VB_Name = "constPython"
Public Sub HighlightPython()
     
    LANGUAGE_ = "Code"
    COMMENT_LINE_ = "#"
    COMMENT_MULTILINE_START_ = ""
    COMMENT_MULTILINE_END_ = ""
    STRING_MULTILINE_START_ = ""
    STRING_MULTILINE_END_ = ""
    RESERVED_ = Array( _
        "and", "as", "assert", "async", "await", "break", "case", "class", "continue", "def", _
        "del", "elif", "else", "except", "finally", "for", "from", "global", "if", "import", _
        "in", "is", "lambda", "match", "nonlocal|10", "not", "or", "pass", "raise", "return", _
        "try", "while", "with", "yield" _
    )
    OPERATORS_ = Array( _
        "==", "!=", ">", "<", ">=", "<=", "+", "-", "*", "/", "//", "%", "**", "=", "&", "|", _
        ">>", "<<", "&&", "||" _
    )
    TYPES_ = Array( _
        "Any", "Callable", "Coroutine", "Dict", "List", "Literal", "Generic", "Optional", "Sequence", "Set", _
        "Tuple", "Type", "Union" _
    )
    BUILTINS_ = Array( _
        "__import__", "abs", "all", "any", "ascii", "bin", "bool", "breakpoint", "bytearray", "bytes", _
        "callable", "chr", "classmethod", "compile", "complex", "delattr", "dict", "dir", "divmod", "enumerate", _
        "eval", "exec", "filter", "float", "format", "frozenset", "getattr", "globals", "hasattr", "hash", _
        "help", "hex", "id", "input", "int", "isinstance", "issubclass", "iter", "len", "list", _
        "locals", "map", "max", "memoryview", "min", "next", "object", "oct", "open", "ord", _
        "pow", "print", "property", "range", "repr", "reversed", "round", "set", "setattr", _
        "slice", "sorted", "staticmethod", "str", "sum", "super", "tuple", "type", "vars", "zip" _
    )
    LITERALS_ = Array( _
        "__debug__", "Ellipsis", "False", "None", "NotImplemented", "True" _
    )
    
End Sub


