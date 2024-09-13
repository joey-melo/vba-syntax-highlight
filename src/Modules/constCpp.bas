Attribute VB_Name = "constCPP"
'@lang VBA

Public Sub HighlightCpp()
     
    LANGUAGE_ = "Code"
    COMMENT_LINE_ = "//"
    COMMENT_MULTILINE_START_ = "/*"
    COMMENT_MULTILINE_END_ = "*/"
    STRING_MULTILINE_START_ = ""
    STRING_MULTILINE_END_ = ""
    RESERVED_ = Array( _
        "alignas", "alignof", "and", "and_eq", "asm", "atomic_cancel", "atomic_commit", "atomic_noexcept", "auto", "bitand", _
        "bitor", "break", "case", "catch", "class", "co_await", "co_return", "co_yield", "compl", "concept", _
        "const_cast|10", "consteval", "constexpr", "constinit", "continue", "decltype", "default", "delete", "do", "dynamic_cast|10", _
        "else", "enum", "explicit", "export", "extern", "false", "final", "for", "friend", "goto", _
        "if", "import", "inline", "module", "mutable", "namespace", "new", "noexcept", "not", "not_eq", _
        "nullptr", "operator", "or", "or_eq", "override", "private", "protected", "public", "reflexpr", "register", _
        "reinterpret_cast|10", "requires", "return", "sizeof", "static_assert", "static_cast|10", "struct", "switch", "synchronized", "template", _
        "this", "thread_local", "throw", "transaction_safe", "transaction_safe_dynamic", "true", "try", "typedef", "typeid", "typename", _
        "union", "using", "virtual", "volatile", "while", "xor", "xor_eq", "#include" _
    )
    OPERATORS_ = Array( _
        "==", "!=", ">", "<", ">=", "<=", "+", "-", "*", "/", "//", "%", "**", "=", "&", "|", _
        ">>", "<<", "&&", "||" _
    )
    TYPES_ = Array( _
        "bool", "char", "char16_t", "char32_t", "char8_t", "double", "float", "int", "long", "short", _
        "void", "wchar_t", "unsigned", "signed", "const", "static" _
    )
    BUILTINS_ = Array( _
        "abort", "abs", "acos", "apply", "as_const", "asin", "atan", "atan2", "calloc", "ceil", _
        "cerr", "cin", "clog", "cos", "cosh", "cout", "declval", "endl", "exchange", "exit", _
        "exp", "fabs", "floor", "fmod", "forward", "fprintf", "fputs", "free", "frexp", "fscanf", _
        "future", "invoke", "isalnum", "isalpha", "iscntrl", "isdigit", "isgraph", "islower", "isprint", "ispunct", _
        "isspace", "isupper", "isxdigit", "labs", "launder", "ldexp", "log", "log10", "make_pair", "make_shared", _
        "make_shared_for_overwrite", "make_tuple", "make_unique", "malloc", "memchr", "memcmp", "memcpy", "memset", "modf", "move", _
        "pow", "printf", "putchar", "puts", "realloc", "scanf", "sin", "sinh", "snprintf", "sprintf", _
        "sqrt", "sscanf", "std", "stderr", "stdin", "stdout", "strcat", "strchr", "strcmp", "strcpy", _
        "strcspn", "strlen", "strncat", "strncmp", "strncpy", "strpbrk", "strrchr", "strspn", "strstr", "swap", _
        "tan", "tanh", "terminate", "to_underlying", "tolower", "toupper", "vfprintf", "visit", "vprintf", "vsprintf", _
         "any", "auto_ptr", "barrier", "binary_semaphore", "bitset", "complex", "condition_variable", "condition_variable_any", "counting_semaphore", "deque", _
        "false_type", "future", "imaginary", "initializer_list", "istringstream", "jthread", "latch", "lock_guard", "multimap", "multiset", _
        "mutex", "optional", "ostringstream", "packaged_task", "pair", "promise", "priority_queue", "queue", "recursive_mutex", "recursive_timed_mutex", _
        "scoped_lock", "set", "shared_future", "shared_lock", "shared_mutex", "shared_timed_mutex", "shared_ptr", "stack", "string_view", "stringstream", _
        "timed_mutex", "thread", "true_type", "tuple", "unique_lock", "unique_ptr", "unordered_map", "unordered_multimap", "unordered_multiset", "unordered_set", _
        "variant", "vector", "weak_ptr", "wstring", "wstring_view" _
    )
    LITERALS_ = Array( _
        "NULL", "true", "false", "nullopt", "nullptr" _
    )
    
End Sub








