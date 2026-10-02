/**
 * Pre-flight scan of an xlsx package.
 *
 * Saving goes through ExcelJS, which rebuilds the package from its own model.
 * Anything it does not model is silently dropped. Rather than let that happen
 * quietly, list the zip entries up front and tell the user exactly which
 * features will not survive a save.
 */
import { unzipSync } from 'fflate';

export interface LossyFeature {
  /** Short Korean label shown in the warning bar. */
  label: string;
  /** How many parts of this kind the file contains. */
  count: number;
}

interface Detector {
  label: string;
  test: (path: string) => boolean;
}

const DETECTORS: Detector[] = [
  { label: '차트', test: (p) => p.startsWith('xl/charts/chart') },
  { label: '피벗 테이블', test: (p) => p.startsWith('xl/pivotTables/') },
  { label: '이미지·도형', test: (p) => p.startsWith('xl/media/') },
  { label: '매크로(VBA)', test: (p) => p.endsWith('vbaProject.bin') },
  { label: '조건부 서식 규칙(확장)', test: (p) => p.startsWith('xl/revisions/') },
  { label: '슬라이서', test: (p) => p.startsWith('xl/slicers/') },
  { label: '표(Table) 서식', test: (p) => p.startsWith('xl/tables/') },
  { label: '주석(메모)', test: (p) => /^xl\/(comments|threadedComments)/.test(p) },
  { label: '외부 링크', test: (p) => p.startsWith('xl/externalLinks/') },
];

/**
 * Lists zip entry names without inflating their contents. fflate's `filter`
 * runs before decompression, so returning `false` for everything turns this
 * into a cheap directory listing.
 */
function listEntries(bytes: Uint8Array): string[] {
  const names: string[] = [];
  unzipSync(bytes, {
    filter: (file) => {
      names.push(file.name);
      return false;
    },
  });
  return names;
}

/** Returns the features present in the file that a save would discard. */
export function detectLossyFeatures(bytes: Uint8Array): LossyFeature[] {
  let entries: string[];
  try {
    entries = listEntries(bytes);
  } catch {
    // A malformed zip is ExcelJS's problem to report, not ours.
    return [];
  }

  const counts = new Map<string, number>();
  for (const path of entries) {
    for (const detector of DETECTORS) {
      if (detector.test(path)) {
        counts.set(detector.label, (counts.get(detector.label) ?? 0) + 1);
      }
    }
  }

  return [...counts.entries()].map(([label, count]) => ({ label, count }));
}

/** One-line summary for the warning bar, or undefined when nothing is lost. */
export function describeLossyFeatures(features: LossyFeature[]): string | undefined {
  if (features.length === 0) return undefined;
  const parts = features.map((f) => (f.count > 1 ? `${f.label} ${f.count}개` : f.label));
  return `이 파일의 ${parts.join(', ')}는 저장할 때 사라집니다.`;
}
