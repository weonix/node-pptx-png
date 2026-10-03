/**
 * ECMA-376 unit conversion constants.
 *
 * EMU (English Metric Unit) is the base unit used in OpenXML.
 * 1 inch = 914400 EMU
 * 1 point = 12700 EMU
 * 1 cm = 360000 EMU
 */
/** EMU per inch */
export const EMU_PER_INCH = 914400;
/** EMU per point (1/72 inch) */
export const EMU_PER_POINT = 12700;
/** EMU per centimeter */
export const EMU_PER_CM = 360000;
/** EMU per pixel at 96 DPI (standard Windows DPI) */
export const EMU_PER_PIXEL_96DPI = EMU_PER_INCH / 96;
/** Angle unit (60,000ths of a degree) per degree */
export const ANGLE_UNIT_PER_DEGREE = 60000;
/** Standard 4:3 slide width in EMU (10 inches) */
export const STANDARD_SLIDE_WIDTH_EMU = 9144000;
/** Standard 4:3 slide height in EMU (7.5 inches) */
export const STANDARD_SLIDE_HEIGHT_EMU = 6858000;
/** Widescreen 16:9 slide width in EMU (13.333 inches) */
export const WIDESCREEN_SLIDE_WIDTH_EMU = 12192000;
/** Widescreen 16:9 slide height in EMU (7.5 inches) */
export const WIDESCREEN_SLIDE_HEIGHT_EMU = 6858000;
/**
 * Unit converter for OpenXML coordinate transformations.
 */
export class UnitConverter {
    dpi;
    emuPerPixel;
    constructor(dpi = 96) {
        this.dpi = dpi;
        this.emuPerPixel = EMU_PER_INCH / dpi;
    }
    /**
     * Converts EMU to pixels at the configured DPI.
     */
    emuToPixels(emu) {
        return emu / this.emuPerPixel;
    }
    /**
     * Converts pixels to EMU at the configured DPI.
     */
    pixelsToEmu(pixels) {
        return pixels * this.emuPerPixel;
    }
    /**
     * Converts EMU to points (1/72 inch).
     */
    emuToPoints(emu) {
        return emu / EMU_PER_POINT;
    }
    /**
     * Converts points to EMU.
     */
    pointsToEmu(points) {
        return points * EMU_PER_POINT;
    }
    /**
     * Converts EMU to inches.
     */
    emuToInches(emu) {
        return emu / EMU_PER_INCH;
    }
    /**
     * Converts inches to EMU.
     */
    inchesToEmu(inches) {
        return inches * EMU_PER_INCH;
    }
    /**
     * Converts EMU to centimeters.
     */
    emuToCm(emu) {
        return emu / EMU_PER_CM;
    }
    /**
     * Converts centimeters to EMU.
     */
    cmToEmu(cm) {
        return cm * EMU_PER_CM;
    }
    /**
     * Converts OpenXML angle (60,000ths of a degree) to radians.
     */
    angleToRadians(angle) {
        return (angle / ANGLE_UNIT_PER_DEGREE) * (Math.PI / 180);
    }
    /**
     * Converts OpenXML angle (60,000ths of a degree) to degrees.
     */
    angleToDegrees(angle) {
        return angle / ANGLE_UNIT_PER_DEGREE;
    }
    /**
     * Converts degrees to OpenXML angle units.
     */
    degreesToAngle(degrees) {
        return degrees * ANGLE_UNIT_PER_DEGREE;
    }
    /**
     * Converts radians to OpenXML angle units.
     */
    radiansToAngle(radians) {
        return ((radians * 180) / Math.PI) * ANGLE_UNIT_PER_DEGREE;
    }
    /**
     * Converts font size in hundredths of a point to points.
     */
    fontSizeToPoints(fontSize) {
        return fontSize / 100;
    }
    /**
     * Converts points to font size in hundredths of a point.
     */
    pointsToFontSize(points) {
        return points * 100;
    }
    /**
     * Converts percentage value (100000 = 100%) to decimal.
     */
    percentageToDecimal(percentage) {
        return percentage / 100000;
    }
    /**
     * Converts decimal to percentage value (100000 = 100%).
     */
    decimalToPercentage(decimal) {
        return decimal * 100000;
    }
    /**
     * Calculates scale factor from EMU dimensions to pixel dimensions.
     */
    calculateScaleFactor(emuWidth, emuHeight, targetWidth, targetHeight) {
        const pixelWidth = this.emuToPixels(emuWidth);
        const pixelHeight = this.emuToPixels(emuHeight);
        const aspectRatio = pixelWidth / pixelHeight;
        let finalWidth;
        let finalHeight;
        if (targetHeight !== undefined) {
            finalWidth = targetWidth;
            finalHeight = targetHeight;
        }
        else {
            finalWidth = targetWidth;
            finalHeight = Math.round(targetWidth / aspectRatio);
        }
        return {
            scaleX: finalWidth / pixelWidth,
            scaleY: finalHeight / pixelHeight,
            width: finalWidth,
            height: finalHeight,
        };
    }
    /**
     * Gets the configured DPI.
     */
    getDpi() {
        return this.dpi;
    }
}
/**
 * Default unit converter instance at 96 DPI.
 */
export const defaultUnitConverter = new UnitConverter(96);
/**
 * Converts EMU to pixels at 96 DPI.
 */
export function emuToPixels(emu) {
    return defaultUnitConverter.emuToPixels(emu);
}
/**
 * Converts EMU to points.
 */
export function emuToPoints(emu) {
    return defaultUnitConverter.emuToPoints(emu);
}
/**
 * Converts points to EMU.
 */
export function pointsToEmu(points) {
    return defaultUnitConverter.pointsToEmu(points);
}
/**
 * Converts OpenXML angle to radians.
 */
export function angleToRadians(angle) {
    return defaultUnitConverter.angleToRadians(angle);
}
/**
 * Converts font size (hundredths of point) to points.
 */
export function fontSizeToPoints(fontSize) {
    return defaultUnitConverter.fontSizeToPoints(fontSize);
}
/**
 * Converts percentage (100000 = 100%) to decimal.
 */
export function percentageToDecimal(percentage) {
    return defaultUnitConverter.percentageToDecimal(percentage);
}
//# sourceMappingURL=UnitConverter.js.map