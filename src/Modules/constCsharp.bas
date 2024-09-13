Attribute VB_Name = "constCsharp"
'@lang VBA

Public Sub HighlightCsharp()
     
    LANGUAGE_ = "Code"
    COMMENT_LINE_ = "//"
    COMMENT_MULTILINE_START_ = "/*"
    COMMENT_MULTILINE_END_ = "*/"
    STRING_MULTILINE_START_ = "@" & Chr(34)
    STRING_MULTILINE_END_ = Chr(34) & "@"
    RESERVED_ = Array( _
        "abstract", "as", "base", "break", "case", "catch", "class", "const", "continue", "do", _
        "else", "event", "explicit", "extern", "finally", "fixed", "for", "foreach", "goto", "if", _
        "implicit", "in", "interface", "internal", "is", "lock", "namespace", "new", "operator", "out", _
        "override", "params", "private", "protected", "public", "readonly", "record", "ref", "return", "scoped", _
        "sealed", "sizeof", "stackalloc", "static", "struct", "switch", "this", "throw", "try", "typeof", _
        "unchecked", "unsafe", "using", "virtual", "void", "volatile", "whil", "add", "alias", "and", _
        "ascending", "args", "async", "await", "by", "descending", "dynamic", "equals", "file", "from", _
        "get", "global", "group", "init", "into", "join", "let", "nameof", "not", "notnull", _
        "on", "or", "orderby", "partial", "record", "remove", "required", "scoped", "select", "set", _
        "unmanaged", "value|0", "var", "when", "where", "with", "yield" _
    )
    OPERATORS_ = Array( _
        "==", "!=", ">", "<", ">=", "<=", "+", "-", "*", "/", "//", "%", "**", "=", "&", "|", _
        ">>", "<<", "&&", "||" _
    )
    TYPES_ = Array( _
        "public", "private", "protected", "static", "internal", "protected", "abstract", "async", "extern", "override", _
        "unsafe", "virtual", "new", "sealed", "partia" _
    )
    BUILTINS_ = Array( _
        "bool", "byte", "char", "decimal", "delegate", "double", "dynamic", "enum", "float", "int", _
        "long", "nint", "nuint", "object", "sbyte", "short", "string", "ulong", "uint", "ushor" _
    )
    LITERALS_ = Array( _
        "default", "null", "true", "false" _
    )
    
End Sub




