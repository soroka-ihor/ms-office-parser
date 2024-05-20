package io.parser.msoffice.model;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.experimental.FieldDefaults;

@Getter
@Setter
@FieldDefaults(makeFinal = true, level = AccessLevel.PRIVATE)
@AllArgsConstructor
public class TextBlockData {
    @JsonProperty("original_text")   String originalText;
    @JsonProperty("translated_text") String translatedText;
    @JsonProperty("font_size")       Double fontSize;
    @JsonProperty("font_color")      String fontColor;
    @JsonProperty("font_color_hex")  String fontColorHex;
    @JsonProperty("font_style")      String fontStyle;
    @JsonProperty("left")            Double left;
    @JsonProperty("top")             Double top;
    @JsonProperty("end_left")        Double endLeft;
    @JsonProperty("end_top")         Double endTop;
}
