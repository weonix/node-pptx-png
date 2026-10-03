"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Logger = void 0;
exports.createLogger = createLogger;
/**
 * Log level priority for filtering.
 */
const LOG_LEVEL_PRIORITY = {
    debug: 0,
    info: 1,
    warn: 2,
    error: 3,
    silent: 4,
};
/**
 * Console-based logger implementation.
 */
class Logger {
    level;
    context;
    levelPriority;
    constructor(level = 'warn', context) {
        this.level = level;
        this.context = context;
        this.levelPriority = LOG_LEVEL_PRIORITY[level];
    }
    /**
     * Creates a child logger with additional context.
     */
    child(context) {
        const fullContext = this.context ? `${this.context}:${context}` : context;
        return new Logger(this.level, fullContext);
    }
    /**
     * Logs a debug message.
     */
    debug(message, data) {
        this.log('debug', message, data);
    }
    /**
     * Logs an info message.
     */
    info(message, data) {
        this.log('info', message, data);
    }
    /**
     * Logs a warning message.
     */
    warn(message, data) {
        this.log('warn', message, data);
    }
    /**
     * Logs an error message.
     */
    error(message, data) {
        this.log('error', message, data);
    }
    /**
     * Internal log method.
     */
    log(level, message, data) {
        if (LOG_LEVEL_PRIORITY[level] < this.levelPriority) {
            return;
        }
        const prefix = this.context ? `[${this.context}]` : '';
        const timestamp = new Date().toISOString();
        const formattedMessage = `${timestamp} ${level.toUpperCase().padEnd(5)} ${prefix} ${message}`;
        switch (level) {
            case 'debug':
                if (data) {
                    console.debug(formattedMessage, data);
                }
                else {
                    console.debug(formattedMessage);
                }
                break;
            case 'info':
                if (data) {
                    console.info(formattedMessage, data);
                }
                else {
                    console.info(formattedMessage);
                }
                break;
            case 'warn':
                if (data) {
                    console.warn(formattedMessage, data);
                }
                else {
                    console.warn(formattedMessage);
                }
                break;
            case 'error':
                if (data) {
                    console.error(formattedMessage, data);
                }
                else {
                    console.error(formattedMessage);
                }
                break;
        }
    }
}
exports.Logger = Logger;
/**
 * Creates a logger instance based on the log level.
 * Note: 'silent' level uses the standard Logger with priority filtering,
 * which filters out all log levels (debug, info, warn, error have lower priority than silent).
 */
function createLogger(level = 'warn', context) {
    return new Logger(level, context);
}
