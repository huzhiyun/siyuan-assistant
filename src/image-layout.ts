import type { ImageExtent } from "./docx";

export interface RasterDimensions {
    sourceWidth: number;
    sourceHeight: number;
    width: number;
    height: number;
}

/**
 * Fit intrinsic pixels to Word's rendered extent before rasterizing a rotation.
 * This preserves the visual aspect ratio without upscaling either intrinsic axis.
 */
export function rotatedRasterDimensions(
    intrinsicWidth: number,
    intrinsicHeight: number,
    rotationDegrees: number,
    extent?: ImageExtent
): RasterDimensions {
    let sourceWidth = intrinsicWidth;
    let sourceHeight = intrinsicHeight;
    if (extent && extent.widthEmu > 0 && extent.heightEmu > 0) {
        const scale = Math.min(intrinsicWidth / extent.widthEmu, intrinsicHeight / extent.heightEmu);
        sourceWidth = Math.max(1, Math.round(extent.widthEmu * scale));
        sourceHeight = Math.max(1, Math.round(extent.heightEmu * scale));
    }
    const swapSides = rotationDegrees === 90 || rotationDegrees === 270;
    return {
        sourceWidth,
        sourceHeight,
        width: swapSides ? sourceHeight : sourceWidth,
        height: swapSides ? sourceWidth : sourceHeight,
    };
}
