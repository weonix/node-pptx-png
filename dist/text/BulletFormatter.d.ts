/**
 * Formats bullet points and numbered lists.
 * Handles different bullet types (char, autoNum, blip) and calculates indentation.
 */
import type { Rgba } from '../types/geometry.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Bullet types supported by PPTX.
 */
export type BulletType = 'none' | 'char' | 'autoNum' | 'blip';
/**
 * Auto-numbering types for numbered lists.
 */
export type AutoNumType = 'arabicPeriod' | 'arabicParenR' | 'arabicParenBoth' | 'arabicPlain' | 'romanUcPeriod' | 'romanLcPeriod' | 'romanUcParenR' | 'romanLcParenR' | 'romanUcParenBoth' | 'romanLcParenBoth' | 'alphaUcPeriod' | 'alphaLcPeriod' | 'alphaUcParenR' | 'alphaLcParenR' | 'alphaUcParenBoth' | 'alphaLcParenBoth' | 'circleNumDbPlain' | 'circleNumWdBlackPlain' | 'circleNumWdWhitePlain';
/**
 * Bullet properties from paragraph properties.
 */
export interface BulletProps {
    /** Bullet type */
    type: BulletType;
    /** Bullet character (for 'char' type) */
    char?: string;
    /** Auto-numbering type (for 'autoNum' type) */
    autoNumType?: AutoNumType;
    /** Starting number for auto-numbering */
    startAt?: number;
    /** Bullet color (undefined = use text color) */
    color?: Rgba;
    /** Bullet size as percentage of text size (100 = same size) */
    sizePercent?: number;
    /** Bullet font family */
    font?: string;
}
/**
 * Formatted bullet result ready for rendering.
 */
export interface FormattedBullet {
    /** Text to render as bullet */
    text: string;
    /** Font family to use for bullet */
    font?: string;
    /** Color to use (undefined = use text color) */
    color?: Rgba;
    /** Size multiplier (1.0 = same as text) */
    sizeMultiplier: number;
    /** Width of the bullet text in current font */
    width?: number;
}
/**
 * Configuration for BulletFormatter.
 */
export interface BulletFormatterConfig {
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Formats bullet points and numbered lists.
 */
export declare class BulletFormatter {
    private readonly logger;
    constructor(config?: BulletFormatterConfig);
    /**
     * Formats a bullet for a paragraph.
     *
     * @param bulletProps Bullet properties from paragraph
     * @param paragraphIndex Index of paragraph within its list context
     * @param level Indentation level (0-8)
     * @returns Formatted bullet ready for rendering, or undefined if no bullet
     */
    formatBullet(bulletProps: BulletProps | undefined, paragraphIndex: number, _level?: number): FormattedBullet | undefined;
    /**
     * Formats a character bullet.
     */
    private formatCharBullet;
    /**
     * Formats an auto-numbered bullet.
     */
    private formatAutoNumBullet;
    /**
     * Formats a number according to the auto-numbering type.
     */
    private formatNumber;
    /**
     * Converts a number to Roman numerals.
     */
    private toRoman;
    /**
     * Converts a number to alphabetic representation (A, B, C, ... AA, AB, ...).
     */
    private toAlpha;
    /**
     * Converts a number to circled number (Unicode).
     */
    private toCircledNumber;
    /**
     * Calculates bullet indentation in EMU based on level.
     *
     * @param level Indentation level (0-8)
     * @param baseIndent Base indentation in EMU (default: 457200 = 0.5 inch)
     * @returns Indentation in EMU
     */
    calculateBulletIndent(level: number, baseIndent?: number): number;
    /**
     * Gets the default bullet character for a given level.
     *
     * @param level Indentation level (0-8)
     * @returns Default bullet character for that level
     */
    getDefaultBulletChar(level: number): string;
}
/**
 * Creates a BulletFormatter instance.
 */
export declare function createBulletFormatter(logger?: ILogger): BulletFormatter;
//# sourceMappingURL=BulletFormatter.d.ts.map