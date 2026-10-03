/**
 * ECMA-376 unit conversion constants.
 *
 * EMU (English Metric Unit) is the base unit used in OpenXML.
 * 1 inch = 914400 EMU
 * 1 point = 12700 EMU
 * 1 cm = 360000 EMU
 */
/** EMU per inch */
export declare const EMU_PER_INCH = 914400;
/** EMU per point (1/72 inch) */
export declare const EMU_PER_POINT = 12700;
/** EMU per centimeter */
export declare const EMU_PER_CM = 360000;
/** EMU per pixel at 96 DPI (standard Windows DPI) */
export declare const EMU_PER_PIXEL_96DPI: number;
/** Angle unit (60,000ths of a degree) per degree */
export declare const ANGLE_UNIT_PER_DEGREE = 60000;
/** Standard 4:3 slide width in EMU (10 inches) */
export declare const STANDARD_SLIDE_WIDTH_EMU = 9144000;
/** Standard 4:3 slide height in EMU (7.5 inches) */
export declare const STANDARD_SLIDE_HEIGHT_EMU = 6858000;
/** Widescreen 16:9 slide width in EMU (13.333 inches) */
export declare const WIDESCREEN_SLIDE_WIDTH_EMU = 12192000;
/** Widescreen 16:9 slide height in EMU (7.5 inches) */
export declare const WIDESCREEN_SLIDE_HEIGHT_EMU = 6858000;
/**
 * Unit converter for OpenXML coordinate transformations.
 */
export declare class UnitConverter {
    private readonly dpi;
    private readonly emuPerPixel;
    constructor(dpi?: number);
    /**
     * Converts EMU to pixels at the configured DPI.
     */
    emuToPixels(emu: number): number;
    /**
     * Converts pixels to EMU at the configured DPI.
     */
    pixelsToEmu(pixels: number): number;
    /**
     * Converts EMU to points (1/72 inch).
     */
    emuToPoints(emu: number): number;
    /**
     * Converts points to EMU.
     */
    pointsToEmu(points: number): number;
    /**
     * Converts EMU to inches.
     */
    emuToInches(emu: number): number;
    /**
     * Converts inches to EMU.
     */
    inchesToEmu(inches: number): number;
    /**
     * Converts EMU to centimeters.
     */
    emuToCm(emu: number): number;
    /**
     * Converts centimeters to EMU.
     */
    cmToEmu(cm: number): number;
    /**
     * Converts OpenXML angle (60,000ths of a degree) to radians.
     */
    angleToRadians(angle: number): number;
    /**
     * Converts OpenXML angle (60,000ths of a degree) to degrees.
     */
    angleToDegrees(angle: number): number;
    /**
     * Converts degrees to OpenXML angle units.
     */
    degreesToAngle(degrees: number): number;
    /**
     * Converts radians to OpenXML angle units.
     */
    radiansToAngle(radians: number): number;
    /**
     * Converts font size in hundredths of a point to points.
     */
    fontSizeToPoints(fontSize: number): number;
    /**
     * Converts points to font size in hundredths of a point.
     */
    pointsToFontSize(points: number): number;
    /**
     * Converts percentage value (100000 = 100%) to decimal.
     */
    percentageToDecimal(percentage: number): number;
    /**
     * Converts decimal to percentage value (100000 = 100%).
     */
    decimalToPercentage(decimal: number): number;
    /**
     * Calculates scale factor from EMU dimensions to pixel dimensions.
     */
    calculateScaleFactor(emuWidth: number, emuHeight: number, targetWidth: number, targetHeight?: number): {
        scaleX: number;
        scaleY: number;
        width: number;
        height: number;
    };
    /**
     * Gets the configured DPI.
     */
    getDpi(): number;
}
/**
 * Default unit converter instance at 96 DPI.
 */
export declare const defaultUnitConverter: UnitConverter;
/**
 * Converts EMU to pixels at 96 DPI.
 */
export declare function emuToPixels(emu: number): number;
/**
 * Converts EMU to points.
 */
export declare function emuToPoints(emu: number): number;
/**
 * Converts points to EMU.
 */
export declare function pointsToEmu(points: number): number;
/**
 * Converts OpenXML angle to radians.
 */
export declare function angleToRadians(angle: number): number;
/**
 * Converts font size (hundredths of point) to points.
 */
export declare function fontSizeToPoints(fontSize: number): number;
/**
 * Converts percentage (100000 = 100%) to decimal.
 */
export declare function percentageToDecimal(percentage: number): number;
//# sourceMappingURL=UnitConverter.d.ts.map