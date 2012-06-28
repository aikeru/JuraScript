//Package "assert"
//http://wiki.commonjs.org/wiki/Unit_Testing/1.0

var AssertionError = function (args) {
    if (!(this instanceof AssertionError))
    { return new AssertionError(args); }

    this.message = args.message;
    this.actual = args.actual;
    this.expected = args.expected;
};
AssertionError.prototype = new Error();

exports.AssertionError = AssertionError;
var ok = function (guard, message_opt) {
    if (!!guard) { return true; }
    else {
        throw new AssertionError({ message: message_opt, actual: (!!guard), expected: true });
    }
};
var equal = function (actual, expected, message_opt) {
    if (actual == expected) { return true; }
    else {
        throw new AssertionError({ message: message_opt, actual: actual, expected: expected });
    }
};
var notEqual = function (actual, expected, message_opt) {
    if (actual != expected) { return true; }
    else {
        throw new AssertionError({ message: message_opt, actual: actual, expected: expected });
    }
};
var strictEqual = function (actual, expected, message_opt) {
    if (actual === expected) { return true; }
    else {
        throw new AssertionError({ message: message_opt, actual: actual, expected: expected });
    }
};
var notStrictEqual = function (actual, expected, message_opt) {
    if (actual !== expected) { return true; }
    else {
        throw new AssertionError({ message: message_opt, actual: actual, expected: expected });
    }
};
var throws = function (block, Error_opt, message_opt) {
    try {
        block();
    } catch (err) {
        return true;
    }
    throw new AssertionError({ message: message_opt, actual: undefined, expected: Error_opt });
};
//TODO: Finish this
//var deepEqual = function (actual, expected, message_opt) {
//    if(actual.prototype != expected.prototype)
//    { throw new AssertionError({ message: message_opt, actual: actual.prototype, expected: expected.prototype });
//    for(var x in expected) {
//        //Check that hasOwnProperty is identical
//        if(Object.prototype.hasOwnProperty.call(expected, x) && !Object.prototype.hasOwnProperty.call(actual, x)) {
//            throw new AssertionError({ message: message_opt, actual: actual[x], expected: expected[x] });
//        }
//        //Check that the value of each is equal
//        if(typeof x == "object") {
//            
//        }
//    }
//};