const fs = require('fs');
const path = require('path');

const TARGET_PROPERTIES = [
  'downloadUrl',
  'url',
  'contentUrl',
  'currentVersion',
  'versions',
  'links',
  'href',
  'blob',
  'attachment',
  'storage',
  'file',
  'signedUrl',
  'download',
  'preview',
  'originalFile',
];

const DOWNLOAD_KEY_HINTS = [
  'download',
  'url',
  'href',
  'blob',
  'storage',
  'file',
  'signed',
  'content',
  'preview',
  'attachment',
  'original',
];

const OUTPUT_FILE = path.join(process.cwd(), 'downloads', 'document-details.json');

function isObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function looksLikeHttpUrl(value) {
  return typeof value === 'string' && /^https?:\/\//i.test(value);
}

function looksLikeApiPath(value) {
  return typeof value === 'string' && value.startsWith('/');
}

function keySuggestsDownload(key) {
  const lower = String(key).toLowerCase();
  return DOWNLOAD_KEY_HINTS.some((hint) => lower.includes(hint));
}

function formatValue(value) {
  if (value === null) {
    return 'null';
  }
  if (value === undefined) {
    return 'undefined';
  }
  if (typeof value === 'string') {
    return `"${value}"`;
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (Array.isArray(value)) {
    return `[Array length=${value.length}]`;
  }
  if (isObject(value)) {
    return `{Object keys=${Object.keys(value).join(', ')}}`;
  }
  return String(value);
}

function prettyPrintFields(data, prefix = '') {
  const lines = [];

  if (Array.isArray(data)) {
    data.forEach((item, index) => {
      const itemPath = `${prefix}[${index}]`;
      if (isObject(item) || Array.isArray(item)) {
        lines.push(`${itemPath}: ${formatValue(item)}`);
        lines.push(...prettyPrintFields(item, itemPath));
      } else {
        lines.push(`${itemPath}: ${formatValue(item)}`);
      }
    });
    return lines;
  }

  if (!isObject(data)) {
    return [`${prefix || 'root'}: ${formatValue(data)}`];
  }

  for (const [key, value] of Object.entries(data)) {
    const fieldPath = prefix ? `${prefix}.${key}` : key;
    lines.push(`${fieldPath}: ${formatValue(value)}`);

    if (isObject(value) || Array.isArray(value)) {
      lines.push(...prettyPrintFields(value, fieldPath));
    }
  }

  return lines;
}

function inspectNode(value, fieldPath, findings, parentKey = '') {
  if (value === null || value === undefined) {
    return;
  }

  if (Array.isArray(value)) {
    value.forEach((item, index) => {
      inspectNode(item, `${fieldPath}[${index}]`, findings, parentKey);
    });
    return;
  }

  if (!isObject(value)) {
    if (looksLikeHttpUrl(value)) {
      findings.downloadCandidates.push({
        path: fieldPath,
        value,
        reason: 'HTTP/HTTPS URL string',
      });
    } else if (looksLikeApiPath(value) && (keySuggestsDownload(fieldPath.split('.').pop()) || parentKey === 'links')) {
      findings.apiPathCandidates.push({
        path: fieldPath,
        value,
        reason: parentKey === 'links'
          ? 'Relative API path inside links object'
          : 'Relative API path on download-related field',
      });
    }
    return;
  }

  for (const [key, child] of Object.entries(value)) {
    const childPath = fieldPath ? `${fieldPath}.${key}` : key;
    const isTarget = TARGET_PROPERTIES.includes(key);

    if (isTarget) {
      findings.targetPropertyHits.push({
        path: childPath,
        key,
        value: child,
        formatted: formatValue(child),
      });
    }

    if (typeof child === 'string') {
      if (looksLikeHttpUrl(child)) {
        findings.downloadCandidates.push({
          path: childPath,
          value: child,
          reason: isTarget ? `Target property "${key}" contains HTTP URL` : 'HTTP/HTTPS URL string',
        });
      } else if (looksLikeApiPath(child) && (isTarget || keySuggestsDownload(key) || key === 'links' || parentKey === 'links')) {
        findings.apiPathCandidates.push({
          path: childPath,
          value: child,
          reason: parentKey === 'links' || key === 'links'
            ? 'Relative API path inside links object'
            : `Relative API path on field "${key}"`,
        });
      }
    }

    inspectNode(child, childPath, findings, key);
  }
}

function summarizeFields(data) {
  const summary = [];

  function walk(value, fieldPath) {
    if (Array.isArray(value)) {
      summary.push({
        path: fieldPath,
        type: 'array',
        length: value.length,
        description: `Array with ${value.length} item(s)`,
      });
      value.forEach((item, index) => walk(item, `${fieldPath}[${index}]`));
      return;
    }

    if (!isObject(value)) {
      summary.push({
        path: fieldPath,
        type: typeof value,
        value,
        description: `Primitive ${typeof value}`,
      });
      return;
    }

    for (const [key, child] of Object.entries(value)) {
      const childPath = fieldPath ? `${fieldPath}.${key}` : key;
      const type = Array.isArray(child) ? 'array' : typeof child;

      let description = `${type}`;
      if (type === 'object' && child !== null) {
        description = `object with keys: ${Object.keys(child).join(', ')}`;
      } else if (type === 'array') {
        description = `array length ${child.length}`;
      } else if (type === 'string') {
        description = `string (${child.length} chars)`;
      }

      summary.push({ path: childPath, type, description, sample: formatValue(child) });
      walk(child, childPath);
    }
  }

  walk(data, 'root');
  return summary;
}

function buildConclusion(findings) {
  const hasHttpDownload = findings.downloadCandidates.length > 0;

  if (hasHttpDownload) {
    return {
      hasDownloadInformation: true,
      message: 'The metadata response contains at least one HTTP/HTTPS URL that may be used for download.',
      additionalEndpointNeeded: null,
    };
  }

  const hasUsefulApiPaths = findings.apiPathCandidates.length > 0;

  return {
    hasDownloadInformation: false,
    message:
      'The metadata response does not expose a direct HTTP download URL. It contains document identifiers, folder/project references, and relative API links only.',
    additionalEndpointNeeded: hasUsefulApiPaths
      ? 'A separate Filevine download endpoint is still required (e.g. a batch-download or content endpoint) using documentId from this metadata. Relative links such as links.self may help locate related resources but are not downloadable URLs.'
      : 'A separate Filevine download endpoint is still required to obtain a signed URL or file bytes. This GET /documents/{documentId} response provides metadata only.',
    observedClues: findings.apiPathCandidates.map((item) => ({
      path: item.path,
      value: item.value,
      note: item.reason,
    })),
  };
}

function saveDocumentDetails(metadata) {
  const downloadsDir = path.dirname(OUTPUT_FILE);
  if (!fs.existsSync(downloadsDir)) {
    fs.mkdirSync(downloadsDir, { recursive: true });
  }

  fs.writeFileSync(OUTPUT_FILE, JSON.stringify(metadata, null, 2), 'utf8');
  return OUTPUT_FILE;
}

function inspectDocumentResponse(metadata) {
  const prettyLines = prettyPrintFields(metadata);
  const findings = {
    targetPropertyHits: [],
    downloadCandidates: [],
    apiPathCandidates: [],
  };

  inspectNode(metadata, '', findings, '');

  const fieldSummary = summarizeFields(metadata);
  const conclusion = buildConclusion(findings);
  const savedTo = saveDocumentDetails(metadata);

  return {
    metadata,
    prettyLines,
    findings,
    fieldSummary,
    conclusion,
    savedTo,
  };
}

function logInspectionReport(report) {
  console.log('\n========== DOCUMENT METADATA (COMPLETE JSON) ==========');
  console.log(JSON.stringify(report.metadata, null, 2));
  console.log('=======================================================\n');

  console.log('========== PRETTY-PRINTED FIELDS ==========');
  report.prettyLines.forEach((line) => console.log(line));
  console.log('===========================================\n');

  console.log('========== TARGET PROPERTY INSPECTION ==========');
  if (report.findings.targetPropertyHits.length === 0) {
    console.log('No top-level target properties found by exact name match during recursive scan.');
  } else {
    report.findings.targetPropertyHits.forEach((hit) => {
      console.log(`${hit.path}: ${hit.formatted}`);
    });
  }
  console.log('================================================\n');

  console.log('========== DOWNLOAD URL CANDIDATES ==========');
  if (report.findings.downloadCandidates.length === 0) {
    console.log('No HTTP/HTTPS download URLs found in the response.');
  } else {
    report.findings.downloadCandidates.forEach((candidate) => {
      console.log(`>>> DOWNLOAD CANDIDATE at ${candidate.path}`);
      console.log(`    Reason: ${candidate.reason}`);
      console.log(`    URL: ${candidate.value}`);
    });
  }
  console.log('=============================================\n');

  console.log('========== API PATH CLUES ==========');
  if (report.findings.apiPathCandidates.length === 0) {
    console.log('No relative API path clues found on download-related fields.');
  } else {
    report.findings.apiPathCandidates.forEach((candidate) => {
      console.log(`>>> API PATH at ${candidate.path}: ${candidate.value}`);
      console.log(`    Reason: ${candidate.reason}`);
    });
  }
  console.log('====================================\n');

  console.log(`Saved full response to: ${report.savedTo}`);
  console.log('\n========== CONCLUSION ==========');
  console.log(report.conclusion.message);
  if (report.conclusion.additionalEndpointNeeded) {
    console.log(`Additional endpoint needed: ${report.conclusion.additionalEndpointNeeded}`);
  }
  if (report.conclusion.observedClues?.length) {
    console.log('Observed clues:');
    report.conclusion.observedClues.forEach((clue) => {
      console.log(`  - ${clue.path}: ${clue.value} (${clue.note})`);
    });
  }
  console.log('================================\n');
}

module.exports = {
  inspectDocumentResponse,
  logInspectionReport,
  OUTPUT_FILE,
};
