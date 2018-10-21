var bleh = (function () {
  'use strict';

  var languages = {
    'en-US': {
      // value for true
      true: 'TRUE',
      // value for false
      false: 'FALSE',
      // separates function arguments
      argumentSeparator: ',',
      // decimal point in numbers
      decimalSeparator: '.',
      // returns number string that can be parsed by Number()
      reformatNumberForJsParsing: function(n) {return n;}
    },
    'de-DE': {
      true: 'WAHR',
      false: 'FALSCH',
      argumentSeparator: ';',
      decimalSeparator: ',',
      reformatNumberForJsParsing: function(n) {
        return n.replace(',', '.');
      }
    }
  };

  var tokenize_1 = tokenize;

  var TOK_TYPE_NOOP      = "noop";
  var TOK_TYPE_OPERAND   = "operand";
  var TOK_TYPE_FUNCTION  = "function";
  var TOK_TYPE_SUBEXPR   = "subexpression";
  var TOK_TYPE_ARGUMENT  = "argument";
  var TOK_TYPE_OP_PRE    = "operator-prefix";
  var TOK_TYPE_OP_IN     = "operator-infix";
  var TOK_TYPE_OP_POST   = "operator-postfix";
  var TOK_TYPE_WSPACE    = "white-space";
  var TOK_TYPE_UNKNOWN   = "unknown";

  var TOK_SUBTYPE_START       = "start";
  var TOK_SUBTYPE_STOP        = "stop";

  var TOK_SUBTYPE_TEXT        = "text";
  var TOK_SUBTYPE_NUMBER      = "number";
  var TOK_SUBTYPE_LOGICAL     = "logical";
  var TOK_SUBTYPE_ERROR       = "error";
  var TOK_SUBTYPE_RANGE       = "range";

  var TOK_SUBTYPE_MATH        = "math";
  var TOK_SUBTYPE_CONCAT      = "concatenate";
  var TOK_SUBTYPE_INTERSECT   = "intersect";
  var TOK_SUBTYPE_UNION       = "union";


  function createToken(value, type, subtype = '') {
    return {value, type, subtype};
  }

  class Tokens {
    constructor() {
      this.items = [];
      this.index = -1;
    }

    add(value, type, subtype) {
      const token = createToken(value, type, subtype);
      this.addRef(token);
      return token;
    }

    addRef(token) {
      this.items.push(token);
    }

    reset() {
      this.index = -1;
    }

    BOF() {
      return this.index <= 0;
    }

    EOF() {
      return this.index >= this.items.length - 1;
    }

    moveNext() {
      if (this.EOF()) return false;
      this.index++;
      return true;
    }

    current() {
      if (this.index == -1) return null;
      return this.items[this.index];
    }

    next() {
      if (this.EOF()) return null;
      return this.items[this.index + 1];
    }

    previous() {
      if (this.index < 1) return null;
      return (this.items[this.index - 1]);
    }

    toArray() {
      return this.items;
    }
  }

  class TokenStack {
    constructor() {
      this.items = [];
    }

    push(token) {
      this.items.push(token);
    }

    pop() {
      const token = this.items.pop();
      return createToken("", token.type, TOK_SUBTYPE_STOP);
    }

    token() {
      if (this.items.length > 0) {
        return this.items[this.items.length - 1];
      } else {
        return null;
      }
    }

    value() {
      return this.token() ? this.token().value : '';
    }

    type() {
      return this.token() ? this.token().type : '';
    }

    subtype() {
      return this.token() ? this.token().subtype : '';
    }
  }

  function tokenize(formula, options) {
    options = options || {};
    options.language = options.language || 'en-US';

    var language = languages[options.language];
    if (!language) {
      var msg = 'Unsupported language ' + options.language + '. Expected one of: '
        + Object.keys(languages).sort().join(', ');
      throw new Error(msg);
    }

    var tokens = new Tokens();
    var tokenStack = new TokenStack();

    var offset = 0;

    var currentChar = function() { return formula.substr(offset, 1); };
    var doubleChar  = function() { return formula.substr(offset, 2); };
    var nextChar    = function() { return formula.substr(offset + 1, 1); };
    var EOF         = function() { return (offset >= formula.length); };
    var isPreviousNonDigitBlank = function() {
      var offsetCopy = offset;
      if (offsetCopy == 0) return true;

      while (offsetCopy > 0) {
        if (!/\d/.test(formula[offsetCopy])) {
          return /\s/.test(formula[offsetCopy]);
        }

        offsetCopy -= 1;
      }
      return false;
    };

    var isNextNonDigitTheRangeOperator = function() {
      var offsetCopy = offset;

      while (offsetCopy < formula.length) {
        if (!/\d/.test(formula[offsetCopy])) {
          return /:/.test(formula[offsetCopy]);
        }

        offsetCopy += 1;
      }
      return false;
    };

    var token = "";

    var inString = false;
    var inPath = false;
    var inRange = false;
    var inError = false;
    var inNumeric = false;

    while (formula.length > 0) {
      if (formula.substr(0, 1) == " ") {
        formula = formula.substr(1);
      } else {
        if (formula.substr(0, 1) == "=") {
          formula = formula.substr(1);
        }
        break;
      }
    }

    var regexSN = /^[1-9]{1}(\.[0-9]+)?E{1}$/;

    while (!EOF()) {

      // state-dependent character evaluation (order is important)

      // double-quoted strings
      // embeds are doubled
      // end marks token

      if (inString) {
        if (currentChar() == "\"") {
          if (nextChar() == "\"") {
            token += "\"";
            offset += 1;
          } else {
            inString = false;
            tokens.add(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_TEXT);
            token = "";
          }
        } else {
          token += currentChar();
        }
        offset += 1;
        continue;
      }

      // single-quoted strings (links)
      // embeds are double
      // end does not mark a token

      if (inPath) {
        if (currentChar() == "'") {
          if (nextChar() == "'") {
            token += "'";
            offset += 1;
          } else {
            inPath = false;
          }
        } else {
          token += currentChar();
        }
        offset += 1;
        continue;
      }

      // bracked strings (range offset or linked workbook name)
      // no embeds (changed to "()" by Excel)
      // end does not mark a token

      if (inRange) {
        if (currentChar() == "]") {
          inRange = false;
        }
        token += currentChar();
        offset += 1;
        continue;
      }

      // error values
      // end marks a token, determined from absolute list of values

      if (inError) {
        token += currentChar();
        offset += 1;
        if ((",#NULL!,#DIV/0!,#VALUE!,#REF!,#NAME?,#NUM!,#N/A,").indexOf("," + token + ",") != -1) {
          inError = false;
          tokens.add(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_ERROR);
          token = "";
        }
        continue;
      }

      if (inNumeric) {
        if ([language.decimalSeparator, 'E', '+', '-'].indexOf(currentChar()) != -1 || /\d/.test(currentChar())) {
          inNumeric = true;
          token += currentChar();

          offset += 1;
          continue;
        } else {
          inNumeric = false;
          tokens.add(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_NUMBER);
          token = "";
        }
      }

      // scientific notation check

      if (("+-").indexOf(currentChar()) != -1) {
        if (token.length > 1) {
          if (regexSN.test(token)) {
            token += currentChar();
            offset += 1;
            continue;
          }
        }
      }

      // independent character evaulation (order not important)

      // function, subexpression, array parameters

      if (currentChar() == language.argumentSeparator) {
        if (token.length > 0) {
          tokens.add(token, TOK_TYPE_OPERAND);
          token = "";
        }

        if (tokenStack.type() == TOK_TYPE_FUNCTION) {
          tokens.add(",", TOK_TYPE_ARGUMENT);

          offset += 1;
          continue;
        }
      }

      if (currentChar() == ",") {
        if (token.length > 0) {
          tokens.add(token, TOK_TYPE_OPERAND);
          token = "";
        }

        tokens.add(currentChar(), TOK_TYPE_OP_IN, TOK_SUBTYPE_UNION);

        offset += 1;
        continue;
      }

      // establish state-dependent character evaluations

      if (/\d/.test(currentChar()) && isPreviousNonDigitBlank() && !isNextNonDigitTheRangeOperator()) {
        inNumeric = true;
        token += currentChar();
        offset += 1;
        continue;
      }

      if (currentChar() == "\"") {
        if (token.length > 0) {
          // not expected
          tokens.add(token, TOK_TYPE_UNKNOWN);
          token = "";
        }
        inString = true;
        offset += 1;
        continue;
      }

      if (currentChar() == "'") {
        if (token.length > 0) {
          // not expected
          tokens.add(token, TOK_TYPE_UNKNOWN);
          token = "";
        }
        inPath = true;
        offset += 1;
        continue;
      }

      if (currentChar() == "[") {
        inRange = true;
        token += currentChar();
        offset += 1;
        continue;
      }

      if (currentChar() == "#") {
        if (token.length > 0) {
          // not expected
          tokens.add(token, TOK_TYPE_UNKNOWN);
          token = "";
        }
        inError = true;
        token += currentChar();
        offset += 1;
        continue;
      }

      // mark start and end of arrays and array rows

      if (currentChar() == "{") {
        if (token.length > 0) {
          // not expected
          tokens.add(token, TOK_TYPE_UNKNOWN);
          token = "";
        }
        tokenStack.push(tokens.add("ARRAY", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
        tokenStack.push(tokens.add("ARRAYROW", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
        offset += 1;
        continue;
      }

      if (currentChar() == ";") {
        if (token.length > 0) {
          tokens.add(token, TOK_TYPE_OPERAND);
          token = "";
        }
        tokens.addRef(tokenStack.pop());
        tokens.add(",", TOK_TYPE_ARGUMENT);
        tokenStack.push(tokens.add("ARRAYROW", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
        offset += 1;
        continue;
      }

      if (currentChar() == "}") {
        if (token.length > 0) {
          tokens.add(token, TOK_TYPE_OPERAND);
          token = "";
        }
        tokens.addRef(tokenStack.pop());
        tokens.addRef(tokenStack.pop());
        offset += 1;
        continue;
      }

      // trim white-space

      if (currentChar() == " ") {
        if (token.length > 0) {
          tokens.add(token, TOK_TYPE_OPERAND);
          token = "";
        }
        tokens.add(currentChar(), TOK_TYPE_WSPACE);
        offset += 1;
        while ((currentChar() == " ") && (!EOF())) {
          offset += 1;
        }
        continue;
      }

      // multi-character comparators

      if ((",>=,<=,<>,").indexOf("," + doubleChar() + ",") != -1) {
        if (token.length > 0) {
          tokens.add(token, TOK_TYPE_OPERAND);
          token = "";
        }
        tokens.add(doubleChar(), TOK_TYPE_OP_IN, TOK_SUBTYPE_LOGICAL);
        offset += 2;
        continue;
      }

      // standard infix operators

      if (("+-*/^&=><").indexOf(currentChar()) != -1) {
        if (token.length > 0) {
          tokens.add(token, TOK_TYPE_OPERAND);
          token = "";
        }
        tokens.add(currentChar(), TOK_TYPE_OP_IN);
        offset += 1;
        continue;
      }

      // standard postfix operators

      if (("%").indexOf(currentChar()) != -1) {
        if (token.length > 0) {
          tokens.add(token, TOK_TYPE_OPERAND);
          token = "";
        }
        tokens.add(currentChar(), TOK_TYPE_OP_POST);
        offset += 1;
        continue;
      }

      // start subexpression or function

      if (currentChar() == "(") {
        if (token.length > 0) {
          tokenStack.push(tokens.add(token, TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
          token = "";
        } else {
          tokenStack.push(tokens.add("", TOK_TYPE_SUBEXPR, TOK_SUBTYPE_START));
        }
        offset += 1;
        continue;
      }

      // stop subexpression

      if (currentChar() == ")") {
        if (token.length > 0) {
          tokens.add(token, TOK_TYPE_OPERAND);
          token = "";
        }
        tokens.addRef(tokenStack.pop());
        offset += 1;
        continue;
      }

      // token accumulation

      token += currentChar();
      offset += 1;

    }

    // dump remaining accumulation

    if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);

    // move all tokens to a new collection, excluding all unnecessary white-space tokens

    var tokens2 = new Tokens();

    while (tokens.moveNext()) {

      token = tokens.current();

      if (token.type == TOK_TYPE_WSPACE) {
        if ((tokens.BOF()) || (tokens.EOF())) ; else if (!(
                   ((tokens.previous().type == TOK_TYPE_FUNCTION) && (tokens.previous().subtype == TOK_SUBTYPE_STOP)) ||
                   ((tokens.previous().type == TOK_TYPE_SUBEXPR) && (tokens.previous().subtype == TOK_SUBTYPE_STOP)) ||
                   (tokens.previous().type == TOK_TYPE_OPERAND)
                  )
                ) ;
        else if (!(
                   ((tokens.next().type == TOK_TYPE_FUNCTION) && (tokens.next().subtype == TOK_SUBTYPE_START)) ||
                   ((tokens.next().type == TOK_TYPE_SUBEXPR) && (tokens.next().subtype == TOK_SUBTYPE_START)) ||
                   (tokens.next().type == TOK_TYPE_OPERAND)
                   )
                 ) ;
        else {
          tokens2.add(token.value, TOK_TYPE_OP_IN, TOK_SUBTYPE_INTERSECT);
        }
        continue;
      }

      tokens2.addRef(token);

    }

    // switch infix "-" operator to prefix when appropriate, switch infix "+" operator to noop when appropriate, identify operand
    // and infix-operator subtypes, pull "@" from in front of function names

    while (tokens2.moveNext()) {

      token = tokens2.current();

      if ((token.type == TOK_TYPE_OP_IN) && (token.value == "-")) {
        if (tokens2.BOF()) {
          token.type = TOK_TYPE_OP_PRE;
        } else if (
                 ((tokens2.previous().type == TOK_TYPE_FUNCTION) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP)) ||
                 ((tokens2.previous().type == TOK_TYPE_SUBEXPR) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP)) ||
                 (tokens2.previous().type == TOK_TYPE_OP_POST) ||
                 (tokens2.previous().type == TOK_TYPE_OPERAND)
               ) {
          token.subtype = TOK_SUBTYPE_MATH;
        } else {
          token.type = TOK_TYPE_OP_PRE;
        }
        continue;
      }

      if ((token.type == TOK_TYPE_OP_IN) && (token.value == "+")) {
        if (tokens2.BOF()) {
          token.type = TOK_TYPE_NOOP;
        } else if (
                 ((tokens2.previous().type == TOK_TYPE_FUNCTION) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP)) ||
                 ((tokens2.previous().type == TOK_TYPE_SUBEXPR) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP)) ||
                 (tokens2.previous().type == TOK_TYPE_OP_POST) ||
                 (tokens2.previous().type == TOK_TYPE_OPERAND)
               ) {
          token.subtype = TOK_SUBTYPE_MATH;
        } else {
          token.type = TOK_TYPE_NOOP;
        }
        continue;
      }

      if ((token.type == TOK_TYPE_OP_IN) && (token.subtype.length == 0)) {
        if (("<>=").indexOf(token.value.substr(0, 1)) != -1) {
          token.subtype = TOK_SUBTYPE_LOGICAL;
        } else if (token.value == "&") {
          token.subtype = TOK_SUBTYPE_CONCAT;
        } else {
          token.subtype = TOK_SUBTYPE_MATH;
        }
        continue;
      }

      if ((token.type == TOK_TYPE_OPERAND) && (token.subtype.length == 0)) {
        if (isNaN(Number(language.reformatNumberForJsParsing(token.value)))) {
          if (token.value == language.true) {
            token.subtype = TOK_SUBTYPE_LOGICAL;
            token.value = 'TRUE';
          } else if (token.value == language.false) {
            token.subtype = TOK_SUBTYPE_LOGICAL;
            token.value = 'FALSE';
          } else {
            token.subtype = TOK_SUBTYPE_RANGE;
          }
        } else {
          token.subtype = TOK_SUBTYPE_NUMBER;
          token.value = language.reformatNumberForJsParsing(token.value);
        }
        continue;
      }

      if (token.type == TOK_TYPE_FUNCTION) {
        if (token.value.substr(0, 1) == "@") {
          token.value = token.value.substr(1);
        }
        continue;
      }

    }

    tokens2.reset();

    // move all tokens to a new collection, excluding all noops

    tokens = new Tokens();

    while (tokens2.moveNext()) {
      if (tokens2.current().type != TOK_TYPE_NOOP) {
        tokens.addRef(tokens2.current());
      }
    }

    tokens.reset();

    return tokens.toArray();
  }

  var excelFormulaTokenizer = {
  	tokenize: tokenize_1
  };

  class Operator {
    constructor(symbol, precendence, operandCount = 2, leftAssociative = true) {
      if (operandCount < 1 || operandCount > 2) {
        throw new Error(`operandCount cannot be ${operandCount}, must be 1 or 2`);
      }

      this.symbol = symbol;
      this.precendence = precendence;
      this.operandCount = operandCount;
      this.leftAssociative = leftAssociative;
    }

    isUnary() {
      return this.operandCount === 1;
    }

    isBinary() {
      return this.operandCount === 2;
    }

    evaluatesBefore(other) {
      if (this === Operator.SENTINEL) return false;
      if (other === Operator.SENTINEL) return true;
      if (other.isUnary()) return false;

      if (this.isUnary()) {
        return this.precendence >= other.precendence;
      } else if (this.isBinary()) {
        if (this.precendence === other.precendence) {
          return this.leftAssociative;
        } else {
          return this.precendence > other.precendence;
        }
      }
    }
  }

  // fake operator with lowest precendence
  Operator.SENTINEL = new Operator('S', 0);

  var operator = Operator;

  class Stack {
    constructor() {
      this.items = [];
    }

    push(value) {
      this.items.push(value);
    }

    pop() {
      return this.items.pop();
    }

    top() {
      return this.items[this.items.length - 1];
    }
  }

  var stack = Stack;

  function create() {
    const operands = new stack();
    const operators = new stack();

    operators.push(operator.SENTINEL);

    return {
      operands,
      operators
    };
  }

  function operator$1(symbol, precendence, operandCount, leftAssociative) {
    return new operator(symbol, precendence, operandCount, leftAssociative);
  }

  var create_1 = create;
  var operator_1 = operator$1;
  var SENTINEL = operator.SENTINEL;

  var shuntingYard = {
  	create: create_1,
  	operator: operator_1,
  	SENTINEL: SENTINEL
  };

  var tokenStream = create$1;

  /**
  * @param Object[] tokens - Tokens from excel-formula-tokenizer
  */
  function create$1(tokens) {
    const end = {};
    const arr = [...tokens, end];
    let index = 0;

    return {
      consume() {
        index += 1;
        if (index >= arr.length) {
          throw new Error('Invalid Syntax');
        }
      },
      getNext() {
        return arr[index];
      },
      nextIs(type, subtype) {
        if (this.getNext().type !== type) return false;
        if (subtype && this.getNext().subtype !== subtype) return false;
        return true;
      },
      nextIsOpenParen() {
        return this.nextIs('subexpression', 'start');
      },
      nextIsTerminal() {
        if (this.nextIsNumber()) return true;
        if (this.nextIsText()) return true;
        if (this.nextIsLogical()) return true;
        if (this.nextIsRange()) return true;
        return false;
      },
      nextIsFunctionCall() {
        return this.nextIs('function', 'start');
      },
      nextIsFunctionArgumentSeparator() {
        return this.nextIs('argument');
      },
      nextIsEndOfFunctionCall() {
        return this.nextIs('function', 'stop');
      },
      nextIsBinaryOperator() {
        return this.nextIs('operator-infix');
      },
      nextIsPrefixOperator() {
        return this.nextIs('operator-prefix');
      },
      nextIsPostfixOperator() {
        return this.nextIs('operator-postfix');
      },
      nextIsRange() {
        return this.nextIs('operand', 'range');
      },
      nextIsNumber() {
        return this.nextIs('operand', 'number');
      },
      nextIsText() {
        return this.nextIs('operand', 'text');
      },
      nextIsLogical() {
        return this.nextIs('operand', 'logical');
      },
      pos() {
        return index;
      }
    };
  }

  var nodeBuilder = {
    functionCall,
    number,
    text,
    logical,
    cell,
    cellRange,
    binaryExpression,
    unaryExpression
  };

  function cell(key, refType) {
    return {
      type: 'cell',
      refType,
      key
    };
  }

  function cellRange(leftCell, rightCell) {
    if (!leftCell) {
      throw new Error('Invalid Syntax');
    }
    if (!rightCell) {
      throw new Error('Invalid Syntax');
    }
    return {
      type: 'cell-range',
      left: leftCell,
      right: rightCell
    };
  }

  function functionCall(name, ...args) {
    const argArray = Array.isArray(args[0]) ? args[0] : args;

    return {
      type: 'function',
      name,
      arguments: argArray
    };
  }

  function number(value) {
    return {
      type: 'number',
      value
    };
  }

  function text(value) {
    return {
      type: 'text',
      value
    };
  }

  function logical(value) {
    return {
      type: 'logical',
      value
    };
  }

  function binaryExpression(operator, left, right) {
    if (!left) {
      throw new Error('Invalid Syntax');
    }
    if (!right) {
      throw new Error('Invalid Syntax');
    }
    return {
      type: 'binary-expression',
      operator,
      left,
      right
    };
  }

  function unaryExpression(operator, expression) {
    if (!expression) {
      throw new Error('Invalid Syntax');
    }
    return {
      type: 'unary-expression',
      operator,
      operand: expression
    };
  }

  // https://www.engr.mun.ca/~theo/Misc/exp_parsing.htm

  const {
    create: createShuntingYard,
    operator: createOperator,
    SENTINEL: SENTINEL$1
  } = shuntingYard;



  var buildTree = parseFormula;

  function parseFormula(tokens) {
    const stream = tokenStream(tokens);
    const shuntingYard$$1 = createShuntingYard();

    parseExpression(stream, shuntingYard$$1);

    const retVal = shuntingYard$$1.operands.top();
    if (!retVal) {
      throw new Error('Syntax error');
    }
    return retVal;
  }

  function parseExpression(stream, shuntingYard$$1) {
    parseOperandExpression(stream, shuntingYard$$1);

    let pos;
    while (true) {
      if (!stream.nextIsBinaryOperator()) {
        break;
      }
      if (pos === stream.pos()) {
        throw new Error('Invalid syntax!');
      }
      pos = stream.pos();
      pushOperator(createBinaryOperator(stream.getNext().value), shuntingYard$$1);
      stream.consume();
      parseOperandExpression(stream, shuntingYard$$1);
    }

    while (shuntingYard$$1.operators.top() !== SENTINEL$1) {
      popOperator(shuntingYard$$1);
    }
  }

  function parseOperandExpression(stream, shuntingYard$$1) {
    if (stream.nextIsTerminal()) {
      shuntingYard$$1.operands.push(parseTerminal(stream));
      // parseTerminal already consumes once so don't need to consume on line below
      // stream.consume()
    } else if (stream.nextIsOpenParen()) {
      stream.consume(); // open paren
      withinSentinel(shuntingYard$$1, function () {
        parseExpression(stream, shuntingYard$$1);
      });
      stream.consume(); // close paren
    } else if (stream.nextIsPrefixOperator()) {
      let unaryOperator = createUnaryOperator(stream.getNext().value);
      pushOperator(unaryOperator, shuntingYard$$1);
      stream.consume();
      parseOperandExpression(stream, shuntingYard$$1);
    } else if (stream.nextIsFunctionCall()) {
      parseFunctionCall(stream, shuntingYard$$1);
    }
  }

  function parseFunctionCall(stream, shuntingYard$$1) {
    const name = stream.getNext().value;
    stream.consume(); // consume start of function call

    const args = parseFunctionArgList(stream, shuntingYard$$1);
    shuntingYard$$1.operands.push(nodeBuilder.functionCall(name, args));

    stream.consume(); // consume end of function call
  }

  function parseFunctionArgList(stream, shuntingYard$$1) {
    const reverseArgs = [];

    withinSentinel(shuntingYard$$1, function () {
      let arity = 0;
      let pos;
      while (true) {
        if (stream.nextIsEndOfFunctionCall())
          break;
        if (pos === stream.pos()) {
          throw new Error('Invalid syntax');
        }
        pos = stream.pos();
        parseExpression(stream, shuntingYard$$1);
        arity += 1;

        if (stream.nextIsFunctionArgumentSeparator()) {
          stream.consume();
        }
      }

      for (let i = 0; i < arity; i++) {
        reverseArgs.push(shuntingYard$$1.operands.pop());
      }
    });

    return reverseArgs.reverse();
  }

  function withinSentinel(shuntingYard$$1, fn) {
    shuntingYard$$1.operators.push(SENTINEL$1);
    fn();
    shuntingYard$$1.operators.pop();
  }

  function pushOperator(operator, shuntingYard$$1) {
    while (shuntingYard$$1.operators.top().evaluatesBefore(operator)) {
      popOperator(shuntingYard$$1);
    }
    shuntingYard$$1.operators.push(operator);
  }

  function popOperator({operators, operands}) {
    if (operators.top().isBinary()) {
      const right = operands.pop();
      const left = operands.pop();
      const operator = operators.pop();
      operands.push(nodeBuilder.binaryExpression(operator.symbol, left, right));
    } else if (operators.top().isUnary()) {
      const operand = operands.pop();
      const operator = operators.pop();
      operands.push(nodeBuilder.unaryExpression(operator.symbol, operand));
    }
  }

  function parseTerminal(stream) {
    if (stream.nextIsNumber()) {
      return parseNumber(stream);
    }

    if (stream.nextIsText()) {
      return parseText(stream);
    }

    if (stream.nextIsLogical()) {
      return parseLogical(stream);
    }

    if (stream.nextIsRange()) {
      return parseRange(stream);
    }
  }

  function parseRange(stream) {
    const next = stream.getNext();
    stream.consume();
    return createCellRange(next.value);
  }

  function createCellRange(value) {
    const parts = value.split(':');

    if (parts.length == 2) {
      return nodeBuilder.cellRange(
        nodeBuilder.cell(parts[0], cellRefType(parts[0])),
        nodeBuilder.cell(parts[1], cellRefType(parts[1]))
      );
    } else {
      return nodeBuilder.cell(value, cellRefType(value));
    }
  }

  function cellRefType(key) {
    if (/^\$[A-Z]+\$\d+$/.test(key)) return 'absolute';
    if (/^\$[A-Z]+$/     .test(key)) return 'absolute';
    if (/^\$\d+$/        .test(key)) return 'absolute';
    if (/^\$[A-Z]+\d+$/  .test(key)) return 'mixed';
    if (/^[A-Z]+\$\d+$/  .test(key)) return 'mixed';
    if (/^[A-Z]+\d+$/    .test(key)) return 'relative';
    if (/^\d+$/          .test(key)) return 'relative';
    if (/^[A-Z]+$/       .test(key)) return 'relative';
  }

  function parseText(stream) {
    const next = stream.getNext();
    stream.consume();
    return nodeBuilder.text(next.value);
  }

  function parseLogical(stream) {
    const next = stream.getNext();
    stream.consume();
    return nodeBuilder.logical(next.value === 'TRUE');
  }

  function parseNumber(stream) {
    let value = Number(stream.getNext().value);
    stream.consume();

    if (stream.nextIsPostfixOperator()) {
      value *= 0.01;
      stream.consume();
    }

    return nodeBuilder.number(value);
  }

  function createUnaryOperator(symbol) {
    const precendence = {
      // negation
      '-': 7
    }[symbol];

    return createOperator(symbol, precendence, 1, true);
  }

  function createBinaryOperator(symbol) {
    const precendence = {
      // cell range union and intersect
      ' ': 8,
      ',': 8,
      // raise to power
      '^': 5,
      // multiply, divide
      '*': 4,
      '/': 4,
      // add, subtract
      '+': 3,
      '-': 3,
      // string concat
      '&': 2,
      // comparison
      '=': 1,
      '<>': 1,
      '<=': 1,
      '>=': 1,
      '>': 1,
      '<': 1
    }[symbol];

    return createOperator(symbol, precendence, 2, true);
  }

  var visit_1 = visit;

  function visit(node, visitor) {
    visitNode(node, visitor);
  }

  function visitNode(node, visitor) {
    switch (node.type) {
      case 'cell':
        visitCell(node, visitor);
        break;
      case 'cell-range':
        visitCellRange(node, visitor);
        break;
      case 'function':
        visitFunction(node, visitor);
        break;
      case 'number':
        visitNumber(node, visitor);
        break;
      case 'text':
        visitText(node, visitor);
        break;
      case 'logical':
        visitLogical(node, visitor);
        break;
      case 'binary-expression':
        visitBinaryExpression(node, visitor);
        break;
      case 'unary-expression':
        visitUnaryExpression(node, visitor);
        break;
    }
  }

  function visitCell(node, visitor) {
    if (visitor.enterCell) visitor.enterCell(node);
    if (visitor.exitCell) visitor.exitCell(node);
  }

  function visitCellRange(node, visitor) {
    if (visitor.enterCellRange) visitor.enterCellRange(node);

    visitNode(node.left, visitor);
    visitNode(node.right, visitor);

    if (visitor.exitCellRange) visitor.exitCellRange(node);
  }

  function visitFunction(node, visitor) {
    if (visitor.enterFunction) visitor.enterFunction(node);

    node.arguments.forEach(arg => visitNode(arg, visitor));

    if (visitor.exitFunction) visitor.exitFunction(node);
  }

  function visitNumber(node, visitor) {
    if (visitor.enterNumber) visitor.enterNumber(node);
    if (visitor.exitNumber) visitor.exitNumber(node);
  }

  function visitText(node, visitor) {
    if (visitor.enterText) visitor.enterText(node);
    if (visitor.exitText) visitor.exitText(node);
  }

  function visitLogical(node, visitor) {
    if (visitor.enterLogical) visitor.enterLogical(node);
    if (visitor.exitLogical) visitor.exitLogical(node);
  }

  function visitBinaryExpression(node, visitor) {
    if (visitor.enterBinaryExpression) visitor.enterBinaryExpression(node);

    visitNode(node.left, visitor);
    visitNode(node.right, visitor);

    if (visitor.exitBinaryExpression) visitor.exitBinaryExpression(node);
  }

  function visitUnaryExpression(node, visitor) {
    if (visitor.enterUnaryExpression) visitor.enterUnaryExpression(node);

    visitNode(node.operand, visitor);

    if (visitor.exitUnaryExpression) visitor.exitUnaryExpression(node);
  }

  var buildTree$1 = buildTree;
  var visit$1 = visit_1;

  var excelFormulaAst = {
  	buildTree: buildTree$1,
  	visit: visit$1
  };

  const {tokenize: tokenize$1} = excelFormulaTokenizer;
  const {buildTree: buildTree$2, visit: visit$2} = excelFormulaAst;

  window.t = (f) => {
  	let tokens = tokenize$1(String(f));
  	let tree = buildTree$2(tokens);
  	console.log(tree);
  };

  const functions = {
  	SUM: (...vals) => vals.reduce((sum, v) => v + sum, 0)
  };

  const getVal = (spreadsheet, tree) => {
  	if (tree.type === 'number' || tree.type === 'text') {
  		return tree.value;
  	}
  	if (tree.type === 'function') {
  		return functions[tree.name](...tree.arguments.map(arg => getVal(spreadsheet, arg)))
  	}
  	if (tree.type === 'cell') {
  		let cell = spreadsheet.cells[tree.key] || { value: '' };
  		return cellContents(spreadsheet, cell);
  	}
  	if (tree.type === 'binary-expression') {
  		if (tree.operator === '+') {
  			return getVal(spreadsheet, tree.left) + getVal(spreadsheet, tree.right);
  		}
  		if (tree.operator === '-') {
  			return getVal(spreadsheet, tree.left) - getVal(spreadsheet, tree.right);
  		}
  		if (tree.operator === '*') {
  			return getVal(spreadsheet, tree.left) * getVal(spreadsheet, tree.right);
  		}
  		if (tree.operator === '/') {
  			return getVal(spreadsheet, tree.left) / getVal(spreadsheet, tree.right);
  		}
  	}
  	console.log(tree);
  };

  const cellContents = (spreadsheet, cell) => {
  	let tokens = tokenize$1(String(cell.value));
  	let tree = buildTree$2(tokens);
  	return getVal(spreadsheet, tree);
  };

  fetch('example.json')
  	.then(result => result.json())
  	.then(spreadsheet => {
  		Object.entries(spreadsheet.cells).forEach(([address, cell]) => {
  			address = address.match(/^([A-Z]+)([0-9]+)$/);
  			cell.col = linum2int(address[1]) - 1;
  			cell.row = Number(address[2]) - 1;
  		});
  		renderSpreadsheet(spreadsheet);
  	});

  const $ = selector => document.querySelector(selector);

  const $canvas = $('canvas');
  const ctx = $canvas.getContext('2d');

  ctx.textAlign = 'left';
  ctx.textBaseline = 'top';
  ctx.translate(0.5, 0.5); // prevents aliasing

  const scrollX = 0;
  const scrollY = 0;

  const renderSpreadsheet = spreadsheet => {
  	Object.values(spreadsheet.cells).forEach(cell => {
  		renderCell(spreadsheet, cell);
  	});
  };

  const CELL_HEIGHT = 20;
  const CELL_WIDTH = 100;
  const STYLE_DEFAULTS = {
  	'background-color': '#FFFFFF',
  	'color': '#000000',
  	'font-weight': 'normal',
  	'font-size': '12px',
  	'font-family': 'Arial',
  	'border-top-width': '1px',
  	'border-top-color': '#ccc',
  	'border-right-width': '1px',
  	'border-right-color': '#ccc',
  	'border-bottom-width': '1px',
  	'border-bottom-color': '#ccc',
  	'border-left-width': '1px',
  	'border-left-color': '#ccc',
  };

  const getStyle = (cell, attr) => cell.style && cell.style[attr] || STYLE_DEFAULTS[attr];

  const renderCell = (spreadsheet, cell) => {
  	let x = cell.col * CELL_WIDTH + scrollX;
  	let y = cell.row * CELL_HEIGHT + scrollY;
  	
  	ctx.fillStyle = getStyle(cell, 'background-color');
  	ctx.fillRect(x, y, CELL_WIDTH, CELL_HEIGHT);


  	ctx.strokeStyle = getStyle(cell, 'border-top-color');
  	ctx.lineWidth = parseInt(getStyle(cell, 'border-top-width'), 10);
  	ctx.beginPath();
  	ctx.moveTo(x, y);
  	ctx.lineTo(x + CELL_WIDTH, y);
  	ctx.stroke();

  	ctx.strokeStyle = getStyle(cell, 'border-right-color');
  	ctx.lineWidth = parseInt(getStyle(cell, 'border-right-width'), 10);
  	ctx.beginPath();
  	ctx.moveTo(x + CELL_WIDTH, y);
  	ctx.lineTo(x + CELL_WIDTH, y + CELL_HEIGHT);
  	ctx.stroke();

  	ctx.strokeStyle = getStyle(cell, 'border-bottom-color');
  	ctx.lineWidth = parseInt(getStyle(cell, 'border-bottom-width'), 10);
  	ctx.beginPath();
  	ctx.moveTo(x + CELL_WIDTH, y + CELL_HEIGHT);
  	ctx.lineTo(x, y + CELL_HEIGHT);
  	ctx.stroke();

  	ctx.strokeStyle = getStyle(cell, 'border-left-color');
  	ctx.lineWidth = parseInt(getStyle(cell, 'border-left-width'), 10);
  	ctx.beginPath();
  	ctx.moveTo(x, y + CELL_HEIGHT);
  	ctx.lineTo(x, y);
  	ctx.stroke();


  	ctx.font = [getStyle(cell, 'font-weight'), getStyle(cell, 'font-size'), getStyle(cell, 'font-family')].join(' ');
  	ctx.fillStyle = getStyle(cell, 'color');
  	ctx.fillText(cellContents(spreadsheet, cell), x, y);
  };

  const linum2int = input => {
  	input = input.replace(/[^A-Za-z]/, '');
  	let output = 0;
  	for (let i = 0; i < input.length; i++) {
  		output = output * 26 + parseInt(input.substr(i, 1), 36) - 9;
  	}
  	return output;
  };

  var script = {

  };

  return script;

}());
