/**
 * Formula engine worker.
 *
 * Univer evaluates formulas off the main thread when given a worker; without it
 * a recalculation on a large sheet blocks rendering and typing.
 */
import { createUniver } from '@univerjs/presets';
import { UniverSheetsCoreWorkerPreset } from '@univerjs/preset-sheets-core/worker';

createUniver({
  presets: [UniverSheetsCoreWorkerPreset()],
});
