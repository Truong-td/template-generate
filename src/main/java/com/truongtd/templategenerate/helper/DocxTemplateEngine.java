package com.truongtd.templategenerate.helper;

import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DocxTemplateEngine {

    // {{name}} hoặc {{application.name}}
    private static final Pattern PLACEHOLDER_PATTERN =
            Pattern.compile("\\{\\{(\\w+(?:\\.\\w+)*)}}");

    public String render(String text, Map<String, Object> context) {
        return renderScalars(text, context);
    }

    private String renderScalars(String text, Map<String, Object> context) {
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(text);
        StringBuffer sb = new StringBuffer();

        while (matcher.find()) {
            String key = matcher.group(1);
            Object value = resolveKey(key, context);
            String replacement = value != null ? String.valueOf(value) : "";
            matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
        }

        matcher.appendTail(sb);
        return sb.toString();
    }

    @SuppressWarnings("unchecked")
    private Object resolveKey(String key, Map<String, Object> context) {
        if (!key.contains(".")) {
            return context.get(key);
        }
        String[] parts = key.split("\\.");
        Object current = context;
        for (String part : parts) {
            if (!(current instanceof Map)) return null;
            current = ((Map<String, Object>) current).get(part);
            if (current == null) return null;
        }
        return current;
    }
}
