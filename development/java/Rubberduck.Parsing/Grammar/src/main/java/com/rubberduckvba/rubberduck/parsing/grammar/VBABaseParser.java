package com.rubberduckvba.rubberduck.parsing.grammar;

import org.antlr.v4.runtime.Parser;
import org.antlr.v4.runtime.Token;
import org.antlr.v4.runtime.TokenStream;

import java.util.regex.Pattern;

public abstract class VBABaseParser extends Parser {
    public VBABaseParser(TokenStream input) {
        super(input);
    }

    protected final int TokenTypeAtRelativePosition(int i) {
        return this._input.LA(i);
    }

    protected final Token TokenAtRelativePosition(int i) {
        return this._input.LT(i);
    }
	
    protected final String TextOf(Token token) {
        return token.getText();
    }

    protected final boolean MatchesRegex(String text, String pattern) {
        return Pattern.compile(pattern).matcher(text).matches();
    }

    protected final boolean EqualsStringIgnoringCase(String actual, String expected) {
        return actual.equalsIgnoreCase(expected);
    }

    protected final boolean EqualsStringIgnoringCase(String actual, String... expectedOptions) {
        for (String expected : expectedOptions) {
            if (actual.equalsIgnoreCase(expected)) {
                return true;
            }
        }
        return false;
    }

    protected final boolean EqualsString(String actual, String expected) {
        return actual.equals(expected);
    }

    protected final boolean EqualsString(String actual, String... expectedOptions) {
        for (String expected : expectedOptions) {
            if (actual.equals(expected)) {
                return true;
            }
        }
        return false;
    }

    protected final boolean IsTokenType(int actual, int... expectedOptions) {
        for (int expected : expectedOptions) {
            if (actual == expected) {
                return true;
            }
        }
        return false;
    }
}
