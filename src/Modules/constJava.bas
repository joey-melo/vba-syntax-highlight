Attribute VB_Name = "constJava"
'@lang VBA

Public Sub HighlightJava()
     
    LANGUAGE_ = "Code"
    COMMENT_LINE_ = "//"
    COMMENT_MULTILINE_START_ = "/*"
    COMMENT_MULTILINE_END_ = "*/"
    STRING_MULTILINE_START_ = ""
    STRING_MULTILINE_END_ = ""
    RESERVED_ = Array( _
        "synchronized", "abstract", "private", "var", "static", "if", "const '", "for", "while", "strictfp", _
        "finally", "protected", "import", "native", "final", "void", "enum", "else", "break", "transient", _
        "catch", "instanceof", "volatile", "case", "assert", "package", "default", "public", "try", "switch", _
        "continue", "throws", "protected", "public", "private", "module", "requires", "exports", "do", "sealed", _
        "yield", "permits", "goto", "when" _
    )
    OPERATORS_ = Array( _
        "==", "!=", ">", "<", ">=", "<=", "+", "-", "*", "/", "//", "%", "**", "=", "&", "|", _
        ">>", "<<", "&&", "||" _
    )
    TYPES_ = Array( _
        "char", "boolean", "long", "float", "int", "byte", "short", "double" _
    )
    BUILTINS_ = Array( _
        "super", "this" _
    )
    LITERALS_ = Array( _
        "null", "true", "false" _
    )
    
End Sub


