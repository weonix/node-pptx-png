import type { LogLevel } from '../types/index.js';
/**
 * Log entry with metadata.
 */
export interface LogEntry {
    level: LogLevel;
    message: string;
    context?: string;
    data?: Record<string, unknown>;
    timestamp: Date;
}
/**
 * Logger interface for the rendering pipeline.
 */
export interface ILogger {
    debug(message: string, data?: Record<string, unknown>): void;
    info(message: string, data?: Record<string, unknown>): void;
    warn(message: string, data?: Record<string, unknown>): void;
    error(message: string, data?: Record<string, unknown>): void;
    child(context: string): ILogger;
}
/**
 * Console-based logger implementation.
 */
export declare class Logger implements ILogger {
    private readonly level;
    private readonly context?;
    private readonly levelPriority;
    constructor(level?: LogLevel, context?: string);
    /**
     * Creates a child logger with additional context.
     */
    child(context: string): ILogger;
    /**
     * Logs a debug message.
     */
    debug(message: string, data?: Record<string, unknown>): void;
    /**
     * Logs an info message.
     */
    info(message: string, data?: Record<string, unknown>): void;
    /**
     * Logs a warning message.
     */
    warn(message: string, data?: Record<string, unknown>): void;
    /**
     * Logs an error message.
     */
    error(message: string, data?: Record<string, unknown>): void;
    /**
     * Internal log method.
     */
    private log;
}
/**
 * Creates a logger instance based on the log level.
 * Note: 'silent' level uses the standard Logger with priority filtering,
 * which filters out all log levels (debug, info, warn, error have lower priority than silent).
 */
export declare function createLogger(level?: LogLevel, context?: string): ILogger;
