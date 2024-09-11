Attribute VB_Name = "constC"
Public Sub HighlightC()
     
    LANGUAGE_ = "Code"
    COMMENT_LINE_ = "//"
    COMMENT_MULTILINE_START_ = "/*"
    COMMENT_MULTILINE_END_ = "*/"
    STRING_MULTILINE_START_ = ""
    STRING_MULTILINE_END_ = ""
    RESERVED_ = Array( _
        "asm", "auto", "break", "case", "continue", "default", "do", "else", "enum", "extern", _
        "for", "fortran", "goto", "if", "inline", "register", "restrict", "return", "sizeof", "typeof", _
        "typeof_unqual", "struct", "switch", "typedef", "union", "volatile", "while", "_Alignas", "_Alignof", "_Atomic", _
        "_Generic", "_Noreturn", "_Static_assert", "_Thread_local", "alignas", "alignof", "noreturn", "static_assert", "thread_local", "_Pragma", _
        "#include" _
    )
    OPERATORS_ = Array( _
        "==", "!=", ">", "<", ">=", "<=", "+", "-", "*", "/", "//", "%", "**", "=", "&", "|", _
        ">>", "<<", "&&", "||" _
    )
    TYPES_ = Array( _
        "float", "double", "signed", "unsigned", "int", "short", "long", "char", "void", "_Bool", _
        "_BitInt", "_Complex", "_Imaginary", "_Decimal32", "_Decimal64", "_Decimal96", "_Decimal128", "_Decimal64x", "_Decimal128x", "_Float16", _
        "_Float32", "_Float64", "_Float128", "_Float32x", "_Float64x", "_Float128x", "const", "static", "constexpr", "complex", _
        "bool", "imaginary" _
    )
    BUILTINS_ = Array( _
        "std", "string", "wstring", "cin", "cout", "cerr", "clog", "stdin", "stdout", "stderr", _
        "stringstream", "istringstream", "ostringstream", "auto_ptr", "deque", "list", "queue", "stack", "vector", "map", _
        "set", "pair", "bitset", "multiset", "multimap", "unordered_set", "unordered_map", "unordered_multiset", "unordered_multimap", "priority_queue", _
        "make_pair", "array", "shared_ptr", "abort", "terminate", "abs", "acos", "asin", "atan2", "atan", _
        "calloc", "ceil", "cosh", "cos", "exit", "exp", "fabs", "floor", "fmod", "fprintf", _
        "fputs", "free", "frexp", "fscanf", "future", "isalnum", "isalpha", "iscntrl", "isdigit", "isgraph", _
        "islower", "isprint", "ispunct", "isspace", "isupper", "isxdigit", "tolower", "toupper", "labs", "ldexp", _
        "log10", "log", "malloc", "realloc", "memchr", "memcmp", "memcpy", "memset", "modf", "pow", _
        "printf", "putchar", "puts", "scanf", "sinh", "sin", "snprintf", "sprintf", "sqrt", "sscanf", _
        "strcat", "strchr", "strcmp", "strcpy", "strcspn", "strlen", "strncat", "strncmp", "strncpy", "strpbrk", _
        "strrchr", "strspn", "strstr", "tanh", "tan", "vfprintf", "vprintf", "vsprintf", "endl", "initializer_list", _
        "unique_ptr" _
    )
    LITERALS_ = Array( _
        "NULL", "true", "false" _
    )
    
End Sub
