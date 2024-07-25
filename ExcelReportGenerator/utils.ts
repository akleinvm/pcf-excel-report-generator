export function convertBase64ToArrayBuffer(base64Content: string): ArrayBuffer {
    const binaryString = atob(base64Content);
    const bytes = Uint8Array.from(binaryString, char => char.charCodeAt(0));

    return bytes.buffer
}

export function downloadBlob(fileName: string, blob: Blob): void {
    const blobUrl = URL.createObjectURL(blob);

    const downloadLink = document.createElement('a');
    downloadLink.href = blobUrl;
    downloadLink.download = fileName;

    document.body.appendChild(downloadLink);

    downloadLink.click();

    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(blobUrl)
}