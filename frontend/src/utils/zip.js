import JSZip from "jszip";

/**
 * Unpacks any .zip files in the provided File array and extracts all inner .pdf files.
 * Preserves existing .pdf files as-is.
 */
export async function unpackFiles(inputFiles) {
  const resultFiles = [];
  let unpackedZipCount = 0;
  let totalPdfsFromZips = 0;

  for (const file of inputFiles) {
    const lower = file.name.toLowerCase();
    if (lower.endsWith(".zip")) {
      unpackedZipCount++;
      try {
        const zip = await JSZip.loadAsync(file);
        const entries = Object.entries(zip.files).filter(
          ([name, entry]) =>
            !entry.dir &&
            name.toLowerCase().endsWith(".pdf") &&
            !name.includes("__MACOSX") &&
            !name.split("/").pop().startsWith("._")
        );

        for (const [entryPath, entry] of entries) {
          const fileName = entryPath.split("/").pop();
          const blob = await entry.async("blob");
          const pdfFile = new File([blob], fileName, { type: "application/pdf" });
          resultFiles.push(pdfFile);
          totalPdfsFromZips++;
        }
      } catch (e) {
        console.error(`Failed to unpack zip file ${file.name}:`, e);
      }
    } else if (lower.endsWith(".pdf")) {
      resultFiles.push(file);
    }
  }

  return {
    files: resultFiles,
    unpackedZipCount,
    totalPdfsFromZips,
  };
}
