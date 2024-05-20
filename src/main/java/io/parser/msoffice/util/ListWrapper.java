package io.parser.msoffice.util;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.experimental.FieldDefaults;

import java.util.List;

import static lombok.AccessLevel.PRIVATE;

@Getter
@AllArgsConstructor(staticName = "of")
@FieldDefaults(level = PRIVATE, makeFinal = true)
public class ListWrapper<T> {
    List<T> list;
}