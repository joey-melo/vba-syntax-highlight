Attribute VB_Name = "constShell"
'@lang VBA

Public Sub HighlightShell()

    LANGUAGE_ = "Shell"
    COMMENT_LINE_ = "#"
    COMMENT_MULTILINE_START_ = ""
    COMMENT_MULTILINE_END_ = ""
    STRING_MULTILINE_START_ = "@" & Chr(34)
    STRING_MULTILINE_END_ = Chr(34) & "@"
    RESERVED_ = Array( _
        "if", "then", "else", "elif", "fi", "time", "for", "while", "until", "in", _
        "do", "done", "case", "esac", "coproc", "function", "select" _
    )
    OPERATORS_ = Array( _
        "==", "!=", ">", "<", ">=", "<=", "+", "-", "*", "/", "//", "%", "**", "=", "&&", "||" _
    )
    TYPES_ = Array( _
         "string", "char", "byte", "int", "long", "bool", "decimal", "single", "double", "DateTime", _
         "xml", "array", "hashtable", "void" _
    )
    BUILTINS_ = Array( _
        "alias", "arch", "autoload", "b2sum", "base32", "base64", "basename", "bg", "bind", "bindkey", _
        "break", "builtin", "bye", "caller", "cap", "cat", "cd", "chcon", "chdir", "chgrp", _
        "chmod", "chown", "chroot", "cksum", "clone", "comm", "command", "comparguments", "compcall", "compctl", _
        "compdescribe", "compfiles", "compgroups", "compquote", "comptags", "comptry", "compvalues", "continue", "cp", "csplit", _
        "cut", "date", "dd", "declare", "df", "dir", "dircolors", "dirname", "dirs", "disable", _
        "disown", "du", "echo", "echotc", "echoti", "emulate", "enable", "env", "eval", "exec", _
        "exit", "expand", "export", "expr", "factor", "fc", "fg", "float", "fmt", "fold", _
        "functions", "getcap", "getln", "getopts", "groups", "hash", "head", "help", "history", "hostid", _
        "id", "integer", "jobs", "join", "kill", "let", "limit", "link", "ln", "local", _
        "log", "logname", "logout", "ls", "mapfile", "md5sum", "mkdir", "mkfifo", "mknod", "mktemp", _
        "mv", "nice", "nl", "noglob", "nohup", "nproc", "numfmt", "od", "paste", "pathchk", _
        "pinky", "popd", "pr", "print", "printenv", "printf", "ptx", "pushd", "pushln", "pwd", _
        "read", "readarray", "readlink", "readonly", "realpath", "rehash", "return", "rm", "rmdir", "runcon", _
        "sched", "seq", "setcap", "setopt", "sha1sum", "sha224sum", "sha256sum", "sha384sum", "sha512sum", "shift", _
        "shred", "shuf", "sleep", "sort", "source", "split", "stat", "stdbuf", "stty", "sudo", _
        "sum", "suspend", "sync", "tac", "tail", "tee", "test", "timeout", "times", "touch", _
        "tr", "trap", "truncate", "tsort", "tty", "ttyctl", "type", "typeset", "ulimit", "umask", _
        "unalias", "uname", "unexpand", "unfunction", "unhash", "uniq", "unlimit", "unlink", "unset", "unsetopt", _
        "uptime", "users", "vared", "vdir", "wait", "wc", "whence", "where", "which", "who", _
        "whoami", "yes", "zcompile", "zformat", "zftp", "zle", "zmodload", "zparseopts", "zprof", "zpty", "zregexparse", _
        "zsocket", "zstyle", "ztcp" _
    )
    LITERALS_ = Array( _
        "true", "false" _
    )
       
End Sub




