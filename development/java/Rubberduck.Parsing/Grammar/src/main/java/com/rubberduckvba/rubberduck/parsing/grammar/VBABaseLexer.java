package com.rubberduckvba.rubberduck.parsing.grammar;

import org.antlr.v4.runtime.CharStream;
import org.antlr.v4.runtime.Lexer;

public abstract class VBABaseLexer extends Lexer {
    public VBABaseLexer() {
        super();
    }

    public VBABaseLexer(CharStream input) {
        super(input);
    }

    protected final int CharAtRelativePosition(int i) {
        return this._input.LA(i);
    }
	
    protected final boolean IsChar(int actual, char expected) {
        return (char)actual == expected;
    }

    protected final boolean IsChar(int actual, char... expectedOptions) {
        char actualAsChar = (char)actual;
        for (char expected : expectedOptions) {
            if (actualAsChar == expected) {
                return true;
            }
        }
        return false;
    }
}
