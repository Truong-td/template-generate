package com.truongtd.templategenerate.util;

import java.util.regex.Pattern;

public class StringUtils {

    public static final Pattern BLOCK_START =
            Pattern.compile("\\{\\{\\?(.*?)}}");
    public static final Pattern BLOCK_END =
            Pattern.compile("\\{\\{/(.*?)}}");
    public static final Pattern SCALAR =
            Pattern.compile("\\{\\{([^{}]+)}}");
    public static final Pattern LIST_IN_ROW =
            Pattern.compile("\\{\\{([a-zA-Z0-9_]+)\\.[^}]+}}");
    public static final Pattern LIST_BLOCK_START =
            Pattern.compile("\\{\\{([a-zA-Z0-9_]+)}}");
    // paragraph chỉ chứa 1 scalar: {{avatar}}, {{user.avatar}}, ...
    public static final Pattern IMAGE_ONLY_PLACEHOLDER =
            Pattern.compile("\\{\\{([^{}]+)}}");

    public static final Pattern COND_START = Pattern.compile("\\{\\{\\?([^}]+)}}");
    public static final Pattern COND_END   = Pattern.compile("\\{\\{\\/([^}]+)}}");

    public static final String FILE_EXTENSION_DOCX = ".docx";

    public static final String DEFAULT_FONT = "Times New Roman";

    public static final int DEFAULT_FONT_SIZE = 11;
}
