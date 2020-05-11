
function ExceptionHelper(config) {

    function throwCorsException() {
        throw "Simulated Cors Exception";
    }

    function throwStrictException(value) {
        "use strict";
        function doThrow() {
            throw value;
        }

        doThrow();
    }

    function captureCorsOnError() {
        "use strict";
        function doCapture() {
            // Ignoring any previous handler
            window.onerror = function (message, url, lineNumber, columnNumber, error) {
                appInsights._onerror({
                    message,
                    url,
                    lineNumber,
                    columnNumber,
                    error: error,
                    evt: window.event
                });
    
                return true;
            }
        }

        doCapture();
    }
    return {
        throw: throwPageException,
        capture: capturePageOnError,
        throwCors: throwCorsException,
        throwStrict: throwStrictException,
        captureStrict: captureStrictPageOnError,
        captureCore: captureCorsOnError
    }
}