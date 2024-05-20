package io.parser.msoffice.parser;

import io.parser.msoffice.model.ParsedDocument;

import java.io.File;

public interface Parser {
    ParsedDocument parse(File documentToParsing);
}
