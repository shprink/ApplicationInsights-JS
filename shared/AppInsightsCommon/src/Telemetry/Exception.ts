// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { StackFrame } from '../Interfaces/Contracts/Generated/StackFrame';
import { ExceptionData } from '../Interfaces/Contracts/Generated/ExceptionData';
import { ExceptionDetails } from '../Interfaces/Contracts/Generated/ExceptionDetails';
import { ISerializable } from '../Interfaces/Telemetry/ISerializable';
import { DataSanitizer } from './Common/DataSanitizer';
import { FieldType } from '../Enums';
import { SeverityLevel } from '../Interfaces/Contracts/Generated/SeverityLevel';
import { Util } from '../Util';
import { IDiagnosticLogger, CoreUtils } from '@microsoft/applicationinsights-core-js';
import { IExceptionInternal, IExceptionDetailsInternal, IExceptionStackFrameInternal, IAutoExceptionTelemetry } from '../Interfaces/IExceptionTelemetry';

const NoMethod = "<no_method>";
const c_MaxCallStackLevels = 20;
const c_MaxCharactersOfAnonymousFunction = 50;
const strError = "error";
const strStack = "stack";

interface IStackEntry {
    signature:string, 
    args:Array<any>,
    toString: () => string
};

interface IStackDetails {
    src: string,
    obj: Array<IStackEntry>,
    cur?: Array<IStackEntry>
}

function _isExceptionDetailsInternal(value:any): value is IExceptionDetailsInternal {
    return "hasFullStack" in value && "typeName" in value && "parsedStack" in value;
}

function _isExceptionInternal(value:any): value is IExceptionInternal {
    return ("ver" in value && "exceptions" in value && "properties" in value);
}

function _createStackEntry(signature:string, args:Array<any>): IStackEntry {
    return {
        signature: signature,
        args: args,
        toString: function () { return this.signature; }
    };
}

function _convertStackObj(errorStack:string): IStackDetails {
    var stack:Array<IStackEntry> = [];
    var lines = errorStack.split("\n");
    for (var lp = 0; lp < lines.length; lp++) {
        stack.push(_createStackEntry(lines[lp], []));
    }

    return {
        src: errorStack,
        obj: stack
    };
}

function _getOperaStack(errorMessage:string): IStackDetails {
    var stack:Array<IStackEntry> = [];
    var lines = errorMessage.split("\n");
    for (var lp = 0; lp < lines.length; lp++) {
        var entry = _createStackEntry(lines[lp], []);
        if (lines[lp + 1]) {
            entry.signature += "@" + lines[lp + 1];
            lp++;
        }

        stack.push(entry);
    }

    return {
        src: errorMessage,
        obj: stack
    };
}

function _getStackFromErrorObj(errorObj:any): IStackDetails {
    let details = null;
    if (errorObj) {
        try {
            /* Using bracket notation is support older browsers (IE 7/8 -- dont remember the version) that throw when using dot 
            notation for undefined objects and we don't want to loose the error from being reported */
            if (errorObj[strStack]) {
                // Chrome/Firefox
                details = _convertStackObj(errorObj[strStack]);
            } else if (errorObj[strError]) {
                // Edge error event provides and error object
                if (errorObj[strError][strStack]) {
                    details = _convertStackObj(errorObj[strError][strStack]);
                }
            } else if (errorObj['exception']) {
                // Edge error event provides and error object
                if (errorObj.exception[strStack]) {
                    details = _convertStackObj(errorObj.exception[strStack]);
                }
            } else if (window['opera'] && errorObj['message']) {
                // Opera
                details = _getOperaStack(errorObj.message);
            }
        } catch (e) {
            // something unexpected happened so to avoid failing to report any error lets swallow the exception 
            // and fallback to the callee/caller method
        }
    }

    return details || {
        src: "",
        obj: null
    };
}

function _getCurrentStackTrace() {
    var callStack:Array<IStackEntry> = [];

    try {
        // Skip any callers that we don't care about (extraLevelsToSkip)
        /// <disable>JS2055.DoNotReferenceBannedTerms</disable>
        var callstackFunction = arguments.callee;
        /// <restore>JS2055.DoNotReferenceBannedTerms</restore>

        // Walk up the callstack and add the method name and arguments to the stack array
        var callstackDepth = 0;
        while (callstackFunction && callstackDepth < c_MaxCallStackLevels) {
            let methodSignature = NoMethod;
            try {
                methodSignature = callstackFunction.toString();
            } catch (e) {
            }

            var copyArgs = [];
            var callArgs = callstackFunction['args'] || callstackFunction['arguments'];
            if (callArgs) {
                for (let lp = 0; lp < callArgs.length; lp++) {
                    copyArgs[lp] = callArgs[lp];
                }
            }

            callStack.push(_createStackEntry(methodSignature, copyArgs));
            callstackFunction = callstackFunction.caller;
            callstackDepth++;
        }
    } catch (e) {
        let currentStack:Array<IStackEntry> = _getStackFromErrorObj(e).obj;

        // Log the exception to the stack for better visibility -- this is normally because we tried to walk back through
        // a Javascript method that is set as "use strict" as caller and callee are not supported for these methods.
        callStack.push(_createStackEntry("<" + e.toString() + ">", []));
        if (currentStack && currentStack.slice) {
            callStack = callStack.concat(currentStack.slice(callStack.length - 1));
        }
    }

    return callStack;
}

function _getStackTrace(errorObj:any, extraLevelsToSkip:number): IStackDetails {
    /// <summary>
    /// Get the current call stack as an array of strings, each element being one level of the stack
    /// This also uses the GetStack() callback to see if partners will be providing the stack for us
    /// </summary>
    /// <param name="traceArgumentValues">Whether or not argument values should be added to the callstack. Recommend false in https, true otherwise</param>
    /// <param name="extraLevelsToSkip">[optional]If we are trying to build the callstack, how many levels should we ignore (because they are part of watson or other error handling code</param>
    /// <param name="includeFunctionBody">True if the function body should be included in the stack frames</param>

    // Attempt to derrived the stack from any available errorObj.stack (may not be present)
    var stackDetails = _getStackFromErrorObj(errorObj);
    var callStack:Array<IStackEntry> = _getCurrentStackTrace();

    // try and remove all internal trace entries
    let idx = 0;
    let cnt = 0;
    // Limit the find depth no more than 6 entries from known internal depth
    let maxSkip = extraLevelsToSkip + 6;
    for (let lp = extraLevelsToSkip; lp < maxSkip && lp < callStack.length; lp++) {
        let entry = callStack[lp];
        if (entry.signature.indexOf("trackException") !== -1) {
            extraLevelsToSkip = lp;
            cnt++;
            if (cnt > 2) {
                // Allow at most to embedded trackException calls
                break;
            }
        }
    }

    callStack = callStack.slice(extraLevelsToSkip);
    if (callStack.length) {
        stackDetails.cur = callStack;
    }

    return stackDetails;
}

function _formatStackTrace(stackDetails:IStackDetails, traceArgumentValues?:boolean, includeFunctionBody?:boolean) {
    let stack = "";

    if (stackDetails) {
        if (stackDetails.src) {
            return stackDetails.src;
        }
    
        CoreUtils.arrForEach(stackDetails.cur, (entry) => {
            let methodSignature = NoMethod;
            try {
                methodSignature = entry.toString();
            } catch(e) {
            }

            var argumentValues = "";
            if (traceArgumentValues) {
                var args = entry.args || entry['arguments'];
                if (args) {
                    argumentValues = _expandArguments(methodSignature, args);
                }
            }

            // Remove any double-whitespace, to avoid going over the query string 2kb really fast
            methodSignature = methodSignature.replace(/\s\s+/ig, " ");

            var stackMessage = _getMethodName(entry, includeFunctionBody);
            if (argumentValues) {
                stackMessage += " -- args:[" + argumentValues + "]";
            }

            stack += stackMessage + "\n";
        });
    }

    return stack;
}

function _parseStack(stack:IStackDetails): _StackFrame[] {
    let parsedStack: _StackFrame[];
    let frames = stack.obj || stack.cur;
    if (frames && frames.length > 0) {
        parsedStack = [];
        let level = 0;

        let totalSizeInBytes = 0;

        CoreUtils.arrForEach(frames, (frame) => {
            let theFrame = frame.toString();
            if (_StackFrame.regex.test(theFrame)) {
                const parsedFrame = new _StackFrame(theFrame, level++);
                totalSizeInBytes += parsedFrame.sizeInBytes;
                parsedStack.push(parsedFrame);
            }
        });

        // DP Constraint - exception parsed stack must be < 32KB
        // remove frames from the middle to meet the threshold
        const exceptionParsedStackThreshold = 32 * 1024;
        if (totalSizeInBytes > exceptionParsedStackThreshold) {
            let left = 0;
            let right = parsedStack.length - 1;
            let size = 0;
            let acceptedLeft = left;
            let acceptedRight = right;

            while (left < right) {
                // check size
                const lSize = parsedStack[left].sizeInBytes;
                const rSize = parsedStack[right].sizeInBytes;
                size += lSize + rSize;

                if (size > exceptionParsedStackThreshold) {

                    // remove extra frames from the middle
                    const howMany = acceptedRight - acceptedLeft + 1;
                    parsedStack.splice(acceptedLeft, howMany);
                    break;
                }

                // update pointers
                acceptedLeft = left;
                acceptedRight = right;

                left++;
                right--;
            }
        }
    }

    return parsedStack;
}

function _getErrorType(errorType: any) {
    // Gets the Error Type by passing the constructor (used to get the true type of native error object).
    if (errorType) {
        if (errorType.name) {
            return errorType.name;
        }

        try {
            var funcNameRegex = /function (.{1,})\(/;
            var results = (funcNameRegex).exec((errorType).constructor.toString());
            return (results && results.length > 1) ? results[1] : "";
        } catch (e) {
            // Ignore
        }
    }

    return "";
}

/**
 * Formats the provided errorObj for display and reporting, it may be a String, Object, integer or undefined depending on the browser.
 * @param errorObj The supplied errorObj
 */
export function _formatErrorCode(errorObj:any) {
    if (errorObj) {
        try {
            if (!CoreUtils.isString(errorObj)) {
                var errorType = _getErrorType(errorObj);
                if (JSON && JSON.stringify) {
                    var result = JSON.stringify(errorObj);
                    if (!result || result === "{}") {
                        if (errorObj[strError]) {
                            // Looks like an MS Error Event
                            errorObj = errorObj[strError];
                            errorType = _getErrorType(errorObj);
                        }

                        // stringify exists so lets just convert to JSON for reporting
                        result = JSON.stringify(errorObj);
                        if (!result || result === "{}") {
                            result = errorObj.toString();
                        }
                    }
                } else {
                    result = errorObj.toString() + " - (Missing JSON.stringify)";
                }

                if (result.indexOf(errorType) !== 0) {
                    return errorType + ":" + result;
                }

                return result;
            }
        } catch (e) {
        }
    }

    // Fallback to just letting the object format itself into a string
    return "" + (errorObj || "");
}

function _expandArguments(methodSignature:string, values:Array<any>) {
    /// <summary>
    /// Returns a string of all the argument name+value pairs
    /// </summary>
    /// <param name="methodSignature">The method signature, including argument names. Example: "_expandArguments(methodSignature,values)"</param>
    /// <param name="values">The values of the arguments, in an array.</param>
    /// <returns>A string of all the argument name+value pairs.  Example: Name1=Val1,Name2=Val2,Name3=Val3</returns>

    // Extract just the argument names from the method signiture, and convert these to an array of names
    var argumentNamesStr = methodSignature.substring(methodSignature.indexOf("(") + 1, methodSignature.indexOf(")"));
    var argumentNames = argumentNamesStr ? argumentNamesStr.split(",") : [];
    var argumentNamesLength = argumentNames.length;

    // Go through all of the names we found and format them as "<name>=<value>"
    var formattedValues = [];
    if (values) {
        for (var i = 0; i < argumentNamesLength; i++) {
            // Valid argument names can't contain a ':'
            if (argumentNames[i] && argumentNames[i].indexOf(":") === -1) {
                if (i < values.length) {
                    formattedValues.push(argumentNames[i] + "=" + _formatArgumentValue(values[i]));
                } else {
                    formattedValues.push(argumentNames[i] + "=undefined");
                }
            }
        }

        // show the extra arguments passed
        for (var lp = argumentNamesLength; lp < values.length; lp++) {
            formattedValues.push(_formatArgumentValue(values[lp]));
        }
    }

    // Convert our array back to a comma seperated string of "<name>=<value>,<name>=<value>,..."
    return formattedValues.join(",");
}

function _formatArgumentValue(originalValue:any) {
    /// <summary>
    /// Returns the given value converted to a nicely formatted and shortened string
    /// </summary>
    /// <param name="originalValue">The value of the argument to format</param>
    /// <returns>The given value converted to a nicely formatted and shortened string</returns>

    var argumentType = typeof (originalValue);
    var formattedValue;

    if (originalValue === null) {
        formattedValue = "null";
    } else if (argumentType === "string") {
        // Take the first 10 characters of the string and ellipsis
        formattedValue = "'" + _shortenStringAddEllipsis(originalValue, 13) + "'";
    } else if (argumentType === "function") {
        formattedValue = _getMethodName(originalValue);
    } else if (argumentType === "object") {
        // Bug 761257 (naeims 1/21/09): Somehow calling toString() on originalValue crashes Safari when
        // originalValue is the range object returned after applying justify left/center/right on multiple
        // lines of text in the RTE.
        // We deal with this by not calling toString() on originalValue in Safari, even if it's defined.
        // The effect is that the stack trace will not be as informative and will show the
        // word "object" for any argument of type object, and not its toString().
        // This appears to be fixed in Safari 4.
        try {
            formattedValue = argumentType;
            if (originalValue.toString && !Util.isSafari()) {
                formattedValue = originalValue.toString();
            }
        } catch(e) {
        }
    } else if (argumentType === "boolean" || argumentType === "number") {
        formattedValue = originalValue.toString();
    } else {
        formattedValue = "[" + argumentType + "]";
    }

    return formattedValue;
}

function _shortenStringAddEllipsis(str:string, maxLength:number) {
    /// <summary>
    /// Takes a string and, if it is longer than maxLength, truncates it and adds ellipsis
    /// </summary>
    /// <param name="str" type="String">The string to shorten</param>
    /// <param name="maxLength">The maximum length of the string returned by this function</param>
    /// <returns>A string of length &lt;= maxLength, with "..." appended if there was truncation</returns>

    if (str && str.length > maxLength) {
        str = str.substr(0, maxLength - 3) + "...";
    }

    return str;
}

function _getMethodName(method:IStackEntry, includeFunctionBody?:boolean) {
    /// <summary>
    /// Given a function object, return the name of the function
    /// This also uses the _getMethodName callback to see if the page owner has more data than us
    /// For anonymous functions, take the first 20 characters of the body
    /// </summary>
    /// <param name="method">The function object</param>
    /// <param name="includeFunctionBody" type="Boolean" optional="true">True if the function body should be included.</param>
    /// <returns>The name of the function</returns>

    let methodSignature:string;
    if (method) {
        // Pull out the method signature.
        var methodStr = (method && method.toString) ? method.toString() : NoMethod;
        var functionNameEndIndex = methodStr.indexOf(")") + 1;
        var functionNameStartIndex = methodStr.indexOf(" ") === 8 ? 9 : 0; // If the format is "function <name>(...", leave "function" out of our returned string
        methodSignature = methodStr.substring(functionNameStartIndex, functionNameEndIndex);

        // Anonymous function, include 50 characters of the body so we can identify the function
        if (includeFunctionBody || (methodSignature.indexOf("function") === 0)) {
            var totalLength = functionNameEndIndex + c_MaxCharactersOfAnonymousFunction;
            methodStr = methodStr.replace(/\s\s*/ig, " ");
            methodSignature = _shortenStringAddEllipsis(methodStr, totalLength) + (totalLength < methodStr.length ? "}" : "");
        }
    }

    if (!methodSignature) {
        methodSignature = (method && method.toString) ? method.toString() : NoMethod;
    }

    return methodSignature;
}

const strError = "error";

function _isExceptionDetailsInternal(value:any): value is IExceptionDetailsInternal {
    return "hasFullStack" in value && "typeName" in value;
}

function _isExceptionInternal(value:any): value is IExceptionInternal {
    return ("ver" in value && "exceptions" in value && "properties" in value);
}

function _getErrorType(errorType: any) {
    // Gets the Error Type by passing the constructor (used to get the true type of native error object).
    let typeName = "";
    if (errorType) {
        typeName = errorType.typeName || errorType.name || "";
        if (!typeName) {
            try {
                var funcNameRegex = /function (.{1,})\(/;
                var results = (funcNameRegex).exec((errorType).constructor.toString());
                typeName = (results && results.length > 1) ? results[1] : "";
            } catch (e) {
                // Ignore
            }
        }
    }

    return typeName;
}

export class Exception extends ExceptionData implements ISerializable {

    public static envelopeType = "Microsoft.ApplicationInsights.{0}.Exception";
    public static dataType = "ExceptionData";

    public id?: string;
    public problemGroup?: string;
    public isManual?: boolean;

    public aiDataContract = {
        ver: FieldType.Required,
        exceptions: FieldType.Required,
        severityLevel: FieldType.Default,
        properties: FieldType.Default,
        measurements: FieldType.Default
    }

    /**
     * Constructs a new instance of the ExceptionTelemetry object
     */
    constructor(logger: IDiagnosticLogger, exception: Error | IExceptionInternal | IAutoExceptionTelemetry, properties?: {[key: string]: any}, measurements?: {[key: string]: number}, severityLevel?: SeverityLevel, id?: string) {
        super();

        if (!_isExceptionInternal(exception)) {
            this.exceptions = [new _ExceptionDetails(logger, exception, properties)];
            this.properties = DataSanitizer.sanitizeProperties(logger, properties);
            this.measurements = DataSanitizer.sanitizeMeasurements(logger, measurements);
            if (severityLevel) { this.severityLevel = severityLevel; }
            if (id) { this.id = id; }
        } else {
            this.exceptions = exception.exceptions;
            this.properties = exception.properties;
            this.measurements = exception.measurements;
            if (exception.severityLevel) { this.severityLevel = exception.severityLevel; }
            if (exception.id) { this.id = exception.id; }
            if (exception.problemGroup) { this.problemGroup = exception.problemGroup; }

            // bool/int types, use isNullOrUndefined
            this.ver = 2; // TODO: handle the CS"4.0" ==> breeze 2 conversion in a better way
            if (!CoreUtils.isNullOrUndefined(exception.isManual)) { this.isManual = exception.isManual; }
        } 
    }

    public static CreateFromInterface(logger: IDiagnosticLogger, exception: IExceptionInternal, properties?: any, measurements?: { [key: string]: number }): Exception {
        const exceptions: _ExceptionDetails[] = exception.exceptions
            && CoreUtils.arrMap(exception.exceptions, (ex: IExceptionDetailsInternal) => _ExceptionDetails.CreateFromInterface(logger, ex));
        const exceptionData = new Exception(logger, {...exception, exceptions}, properties, measurements);
        return exceptionData;
    }

    public toInterface(): IExceptionInternal {
        const { exceptions, properties, measurements, severityLevel, ver, problemGroup, id, isManual } = this;

        const exceptionDetailsInterface = exceptions instanceof Array
            && CoreUtils.arrMap(exceptions, (exception: _ExceptionDetails) => exception.toInterface())
            || undefined;

        return {
            ver: "4.0", // TODO: handle the CS"4.0" ==> breeze 2 conversion in a better way
            exceptions: exceptionDetailsInterface,
            severityLevel,
            properties,
            measurements,
            problemGroup,
            id,
            isManual
        } as IExceptionInternal;
    }

    /**
     * Creates a simple exception with 1 stack frame. Useful for manual constracting of exception.
     */
    public static CreateSimpleException(message: string, typeName: string, assembly: string, fileName: string,
        details: string, line: number): Exception {

        return {
            exceptions: [
                {
                    hasFullStack: true,
                    message,
                    stack: details,
                    typeName
                } as ExceptionDetails
            ]
        } as Exception;
    }

    public static formatError = _formatErrorCode;
}

export class _ExceptionDetails extends ExceptionDetails implements ISerializable {

    public aiDataContract = {
        id: FieldType.Default,
        outerId: FieldType.Default,
        typeName: FieldType.Required,
        message: FieldType.Required,
        hasFullStack: FieldType.Default,
        stack: FieldType.Default,
        parsedStack: FieldType.Array
    };

    constructor(logger: IDiagnosticLogger, exception: Error | IExceptionDetailsInternal | IAutoExceptionTelemetry, properties?: {[key: string]: any}) {
        super();

        if (!_isExceptionDetailsInternal(exception)) {
            let error = exception as any;
            if (!Util.isError(error)) {
                error = error[strError] || error.evt || error;
            }
            this.typeName = DataSanitizer.sanitizeString(logger, _getErrorType(error)) || Util.NotSpecified;
            this.message = DataSanitizer.sanitizeMessage(logger, _formatErrorCode(exception.message || error.message || error.description || error)) || Util.NotSpecified;
            const stack = _getStackTrace(error, 6);
            this.parsedStack = _parseStack(stack);
            this[strStack] = DataSanitizer.sanitizeException(logger, _formatStackTrace(stack));
            this.hasFullStack = Util.isArray(this.parsedStack) && this.parsedStack.length > 0;

            if (properties) {
                properties.message = properties.message || this.message;
                properties.errorCodeMsg = properties.errorCodeMsg || _formatErrorCode(exception);
            }
        } else {
            this.typeName = exception.typeName;
            this.message = exception.message;
            this[strStack] = exception[strStack];
            this.parsedStack = exception.parsedStack
            this.hasFullStack = exception.hasFullStack
        }
    }

    public toInterface(): IExceptionDetailsInternal {
        const parsedStack = this.parsedStack instanceof Array
            && CoreUtils.arrMap(this.parsedStack, (frame: _StackFrame) => frame.toInterface());

        const exceptionDetailsInterface: IExceptionDetailsInternal = {
            id: this.id,
            outerId: this.outerId,
            typeName: this.typeName,
            message: this.message,
            hasFullStack: this.hasFullStack,
            stack: this[strStack],
            parsedStack: parsedStack || undefined
        };

        return exceptionDetailsInterface;
    }

    public static CreateFromInterface(logger:IDiagnosticLogger, exception: IExceptionDetailsInternal): _ExceptionDetails {
        const parsedStack = (exception.parsedStack instanceof Array
            &&CoreUtils.arrMap(exception.parsedStack, frame => _StackFrame.CreateFromInterface(frame)))
            || exception.parsedStack;

        const exceptionDetails = new _ExceptionDetails(logger, {...exception, parsedStack});

        return exceptionDetails;
    }
}

export class _StackFrame extends StackFrame implements ISerializable {

    // regex to match stack frames from ie/chrome/ff
    // methodName=$2, fileName=$4, lineNo=$5, column=$6
    public static regex = /^([\s]+at)?([^\@\s\()]*?)(\@|\s\(|\s)([^\(\@\n]+):([0-9]+):([0-9]+)(\)?)$/;
    public static baseSize = 58; // '{"method":"","level":,"assembly":"","fileName":"","line":}'.length
    public sizeInBytes = 0;

    public aiDataContract = {
        level: FieldType.Required,
        method: FieldType.Required,
        assembly: FieldType.Default,
        fileName: FieldType.Default,
        line: FieldType.Default,
    };

    constructor(sourceFrame: string | IExceptionStackFrameInternal, level: number) {
        super();

        // Not converting this to CoreUtils.isString() as typescript uses this logic to "understand" the different
        // types for the 2 different code paths
        if (typeof sourceFrame === "string") {
            const frame: string = sourceFrame;
            this.level = level;
            this.method = NoMethod;
            this.assembly = Util.trim(frame);
            this.fileName = "";
            this.line = 0;
            const matches = frame.match(_StackFrame.regex);
            if (matches && matches.length >= 5) {
                this.method = Util.trim(matches[2]) || this.method;
                this.fileName = Util.trim(matches[4]);
                this.line = parseInt(matches[5]) || 0;
            }
        } else {
            this.level = sourceFrame.level;
            this.method = sourceFrame.method;
            this.assembly = sourceFrame.assembly;
            this.fileName = sourceFrame.fileName;
            this.line = sourceFrame.line;
            this.sizeInBytes = 0;
        }

        this.sizeInBytes += this.method.length;
        this.sizeInBytes += this.fileName.length;
        this.sizeInBytes += this.assembly.length;

        // todo: these might need to be removed depending on how the back-end settles on their size calculation
        this.sizeInBytes += _StackFrame.baseSize;
        this.sizeInBytes += this.level.toString().length;
        this.sizeInBytes += this.line.toString().length;
    }

    public static CreateFromInterface(frame: IExceptionStackFrameInternal) {
        return new _StackFrame(frame, null /* level is available in frame interface */);
    }

    public toInterface() {
        return {
            level: this.level,
            method: this.method,
            assembly: this.assembly,
            fileName: this.fileName,
            line: this.line
        };
    }
}
