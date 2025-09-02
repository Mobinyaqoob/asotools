function isFileTypeSupported(fileType, file) {
  // Check for supported file types

  if (lib.heicTo.isHeic(file) && isHeicExt(file)) {
    fileType = "image/heic";
    ui.outputFileType = "image/heic";
    console.log('File type is HEIC: ', fileType)
  }

  const supportedFileTypes = [
    "image/jpeg",
    "image/png",
    "image/webp",
    "image/heic",
    "image/heif",
    "image/avif",
    "image/gif",
    "image/svg+xml",
    "image/jxl",
    "image/vnd.microsoft.icon",
    "image/x-icon",
    "image/tiff",
  ];

  return supportedFileTypes.includes(fileType);
}

function isPostProcessingRequired(targetOutputfileType) {

  const postProcessingTypes = [
    /**
     * Add mime types which browsers cannot natively output
     * and requires custom encoding to handle the final image file.
     * Implement the handling in `postProcessImage()`.
     */
    "image/vnd.microsoft.icon",
    "image/x-icon",
  ];

  return postProcessingTypes.includes(targetOutputfileType);
}

function mimeToExtension(mimeType) {
  const fileExtensionMap = {
    "image/jpeg": "jpg",
    "image/png": "png",
    "image/webp": "webp",
    "image/heic": "heic",
    "image/heic": "heif",
    "image/avif": "avif",
    "image/gif": "gif",
    "image/svg+xml": "svg",
    "image/jxl": "jxl",
    "image/vnd.microsoft.icon": "ico",
    "image/x-icon": "ico",
    "image/tiff": "tiff",
    "image/dng": "tiff",
    "image/x-adobe-dng": "tiff",
  };

  return (
    fileExtensionMap[mimeType] || mimeType.replace("image/", "").split("+")[0]
  );
}

function defaultConversionMapping(mimeType) {
  const conversionMap = {
    // Image file types that cannot be compressed to its original file format
    // are converted to a relevant counterpart.
    "image/heic": "image/png",
    "image/heif": "image/png",
    "image/avif": "image/png",
    "image/gif": "image/png",
    "image/svg+xml": "image/png",
    "image/vnd.microsoft.icon": "image/png",
    "image/x-icon": "image/png",
  };

  console.log('Input mimeType ', mimeType);
  console.log('Mapped mimeType ', conversionMap[mimeType]);

  return conversionMap[mimeType] || mimeType;
}

function isHeicExt(file) {
  // Checks if file name ending with `.heic` or `.heif`.
  const fileName = file.name.toLowerCase();
  return fileName.endsWith('.heic') || fileName.endsWith('.heif');
}

function isFileExt(file, extension = "") {
  // Checks if file name ending with the passed string argument.
  const fileName = file.name.toLowerCase();
  return fileName.endsWith(`.${extension}`);
}

/**
 * Determines the input and output file types and extensions based on the user-uploaded file
 * and the user-selected target format.
 *
 * @param {File} file - Image file object.
 * @returns {Object} An object containing:
 *   @property {string} inputFileType - The mime type of the uploaded file (e.g., "image/jpeg").
 *   @property {string} inputFileExtension - The file extension of the uploaded file (e.g., "jpg").
 *   @property {string} outputFileExtension - The target file extension after compression (e.g., "webp").
 *   @property {string} selectedFormat - The mime type of the user-selected output format (e.g., "image/webp").
 */
function getFileType(file) {
  let selectedFormat = getCheckedValue(ui.inputs.formatSelect);
  let inputFileType = file.type;
  let inputFileExtension = "";
  let outputFileExtension = "";

  if (selectedFormat && selectedFormat !== "default") {
    // The user-selected format to convert to.
    const extension = mimeToExtension(selectedFormat);
    inputFileExtension = extension;
    outputFileExtension = extension;
  } else {
    // User has not selected a file format, use the input image's file type.
    selectedFormat = file.type || "png";
    file.type = !file.type && isHeicExt(file) ? "image/heic" : file.type;
    inputFileExtension = mimeToExtension(file.type) || "";

    console.log("inputFileExtension: ", inputFileExtension);
    outputFileExtension = mimeToExtension(defaultConversionMapping(file.type));
    console.log("outputFileExtension: ", outputFileExtension);
  }

  return {
    inputFileType,
    inputFileExtension,
    outputFileExtension,
    selectedFormat,
  };
}

function updateFileExtension(originalName, fileExtension, selectedFormat) {
  const baseName = originalName.replace(/\.[^/.]+$/, "");
  const newExtension = selectedFormat
    ? mimeToExtension(fileExtension)
    : fileExtension;

  console.log('New image extension: ', newExtension);
  return `${baseName}.${newExtension}`;
}

function appendFileNameId(fileName = "image") {
  if (typeof fileName !== 'string') return null;

  const lastDotIndex = fileName.lastIndexOf('.');
  const fileExt = (lastDotIndex === -1 || lastDotIndex === 0) ? '' : fileName.slice(lastDotIndex).toLowerCase();
  const baseFileName = (lastDotIndex === -1) ? fileName : fileName.slice(0, lastDotIndex);

  const fileId = Math.random().toString(36).substring(2, 6).toUpperCase();
  return baseFileName + "-" + fileId + fileExt;
}

function renameBrowserDefaultFileName(fileName) {
  /**
   * Naive approach to check if an image was pasted from clipboard and received a default name by the browser,
   * e.g., `image.png`. This method is potentially browser and language-dependent, if naming conventions vary.
   * `HEIF Image.heic` concerns iOS devices, e.g. when drag-and-dropping a subject cut-out.
   */ 
  const defaultNames = [/^image\.\w+$/i, /^heif image\.heic$/i];

  if (defaultNames.some(regex => regex.test(fileName))) {
    return { renamedFileName: appendFileNameId(fileName), isBrowserDefaultFileName: true };
  }
  return { renamedFileName: fileName, isBrowserDefaultFileName: false };
}

function validateWeight(value, unit = "MB") {
  value = Number(value);
  let [min, max] = [config.weightLimit.min, config.weightLimit.max];
  min = unit.toUpperCase() === "KB" ? min * 1000 : min; 
  max = unit.toUpperCase() === "KB" ? max * 1000 : max; 

  if (typeof value !== 'number' || isNaN(value) || !Number.isFinite(value)) {
    const message = "Invalid value, not a number.";
    return {value: null, message}
  }
  else if (value < min) {
    const message = `Minimum file size is ${min * 1000}KB or ${max}MB.`;
    return {value: min, message}
  }
  else if (value > max) {
    const message = `Max file size is ${max}MB.`;
    return {value: max, message}
  }

  return {value, message: null}
}

function getCheckedValue(nodeList) {
  // Find the currently select radio button value.
  return [...nodeList].find((el) => el.checked)?.value || null;
}

function getImageDimensions(imageInput) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    let objectUrl;

    if (imageInput instanceof Blob) {
      objectUrl = URL.createObjectURL(imageInput);
      img.src = objectUrl;
    } else if (typeof imageInput === "string") {
      img.src = imageInput;
    } else {
      reject(new Error("Invalid input provided to getImageDimensions."));
      return;
    }

    img.onload = () => {
      if (objectUrl) URL.revokeObjectURL(objectUrl);
      resolve({
        outputImageWidth: img.naturalWidth,
        outputImageHeight: img.naturalHeight,
      });
    };

    img.onerror = () => {
      if (objectUrl) URL.revokeObjectURL(objectUrl);
      reject(new Error("Failed to load image for dimensions."));
    };
  });
}

function getAdjustedDimensions({ imageBlob, width, height }, desiredLimitDimensions) {
  // Adjusts image dimensions to prevent the short edge from being 0.
  // Calculates the minimum long edge based on a 1px short edge while keeping aspect ratio.

  return new Promise((resolve) => {
    const compute = (w, h) => {
      if (!w || !h) return resolve(undefined);
      const shortEdge = Math.min(w, h);
      const longEdge = Math.max(w, h);
      const minAllowedDimension = Math.ceil(longEdge * (1 / shortEdge));
      resolve(Math.max(desiredLimitDimensions, minAllowedDimension));
    };

    if (typeof width === 'number' && typeof height === 'number') {
      compute(width, height);
    } else if (imageBlob instanceof Blob) {
      getImageDimensions(imageBlob).then(({ outputImageWidth, outputImageHeight }) => {
        compute(outputImageWidth, outputImageHeight);
      });
    } else {
      resolve(undefined);
    }
  });
}


function debugBlobImageOutput(blob) {
  const blobURL = URL.createObjectURL(blob);
  const img = document.createElement("img");
  img.src = blobURL;
  img.style.maxWidth = "100%";
  img.style.display = "block";
  document.body.prepend(img);
}