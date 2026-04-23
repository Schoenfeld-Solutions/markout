export function readDroppedMarkdownFile(file: File): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();

    reader.onerror = () => {
      reject(new Error(`MarkOut could not read ${file.name}.`));
    };
    reader.onload = () => {
      if (typeof reader.result !== "string") {
        reject(new Error(`MarkOut could not decode ${file.name}.`));
        return;
      }

      resolve(reader.result);
    };

    reader.readAsText(file);
  });
}

export function supportsMarkdownFile(file: File): boolean {
  return /\.(md|markdown|txt)$/i.test(file.name);
}
