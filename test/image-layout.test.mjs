import { createRequire } from "module";
const require = createRequire(import.meta.url);
const { rotatedRasterDimensions } = require("/tmp/syass_image_layout.cjs");

let pass = 0;
let fail = 0;
function assert(name, cond, extra = "") {
    if (cond) { pass++; console.log(`  ✓ ${name}`); }
    else { fail++; console.log(`  ✗ ${name} ${extra}`); }
}

const rightAngle = rotatedRasterDimensions(400, 300, 90, { widthEmu: 200, heightEmu: 100 });
assert("wp:extent drives pre-rotation dimensions", rightAngle.sourceWidth === 400 && rightAngle.sourceHeight === 200, JSON.stringify(rightAngle));
assert("90 degree rotation swaps the rendered dimensions", rightAngle.width === 200 && rightAngle.height === 400, JSON.stringify(rightAngle));

const noExtent = rotatedRasterDimensions(400, 300, 270);
assert("no extent preserves intrinsic source dimensions", noExtent.sourceWidth === 400 && noExtent.sourceHeight === 300, JSON.stringify(noExtent));
assert("270 degree rotation swaps intrinsic dimensions", noExtent.width === 300 && noExtent.height === 400, JSON.stringify(noExtent));

console.log(`\n结果: ${pass} 通过 / ${fail} 失败`);
process.exit(fail > 0 ? 1 : 0);
