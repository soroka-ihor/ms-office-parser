package io.parser.msoffice.adviser;

import io.parser.msoffice.parser.Parser;

import java.io.File;

public interface ParserAdviser {
    Parser advise(File file);
}
