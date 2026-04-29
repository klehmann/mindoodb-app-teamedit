/**
 * Helpers that mirror Crepe's image-block resize behaviour in our preview and
 * print views.
 *
 * Crepe stores the user-applied scale as a numeric "ratio" inside the
 * markdown image's `alt` field (e.g. `![0.50](src "caption")`) and applies it
 * by computing the natural-fit height of the image and multiplying by the
 * ratio (see `node_modules/@milkdown/components/lib/image-block/index.js`,
 * `onImageLoad`).
 *
 * `applyImageRatios()` walks an already-rendered tree, looks for
 * `<img data-teamedit-image-ratio="...">` elements, and applies the same
 * height calculation once each image has loaded.  We deliberately mutate
 * `style.height` only and leave `width` to the browser so aspect ratios stay
 * intact, exactly like Crepe does.
 */

const IMAGE_RATIO_ATTR = "data-teamedit-image-ratio";

interface ApplyImageRatiosOptions {
  /**
   * Optional cap that matches the editor's `maxWidth` config.  When provided,
   * the natural-fit width is clamped to this value before the ratio is
   * applied, just like Crepe's `onImageLoad`.
   */
  maxWidth?: number;
  /**
   * Optional cap that matches the editor's `maxHeight` config.  Applied after
   * width clamping but before the ratio multiplier.
   */
  maxHeight?: number;
}

function findContainerWidth(image: HTMLImageElement) {
  // Walk up to the nearest block-level ancestor with a measurable width.  In
  // both the preview pane and the print document the markdown article is the
  // first ancestor with a non-zero width, but we keep the loop generic.
  let current: HTMLElement | null = image.parentElement;
  while (current) {
    const width = current.getBoundingClientRect().width;
    if (width > 0) {
      return width;
    }
    current = current.parentElement;
  }
  return 0;
}

function applyRatioToImage(image: HTMLImageElement, options: ApplyImageRatiosOptions) {
  const ratioAttr = image.getAttribute(IMAGE_RATIO_ATTR);
  if (!ratioAttr) {
    return;
  }
  const ratio = Number.parseFloat(ratioAttr);
  if (!Number.isFinite(ratio) || ratio <= 0) {
    return;
  }
  const naturalWidth = image.naturalWidth;
  const naturalHeight = image.naturalHeight;
  if (!naturalWidth || !naturalHeight) {
    return;
  }

  let containerWidth = findContainerWidth(image);
  if (options.maxWidth && (!containerWidth || options.maxWidth < containerWidth)) {
    containerWidth = options.maxWidth;
  }
  if (!containerWidth) {
    containerWidth = naturalWidth;
  }

  // Mirror Crepe's algorithm: when the image fits inside the container, use
  // its natural height; otherwise scale to fit width while preserving aspect.
  let fittedHeight = naturalWidth < containerWidth
    ? naturalHeight
    : containerWidth * (naturalHeight / naturalWidth);
  if (options.maxHeight && fittedHeight > options.maxHeight) {
    fittedHeight = options.maxHeight;
  }

  const targetHeight = fittedHeight * ratio;
  image.style.height = `${targetHeight.toFixed(2)}px`;
}

function trackImage(image: HTMLImageElement, options: ApplyImageRatiosOptions) {
  if (image.complete && image.naturalWidth > 0) {
    applyRatioToImage(image, options);
    return;
  }
  const onLoad = () => {
    image.removeEventListener("load", onLoad);
    image.removeEventListener("error", onError);
    applyRatioToImage(image, options);
  };
  const onError = () => {
    image.removeEventListener("load", onLoad);
    image.removeEventListener("error", onError);
  };
  image.addEventListener("load", onLoad, { once: true });
  image.addEventListener("error", onError, { once: true });
}

/**
 * Apply Crepe-style ratio sizing to every image inside `root` that carries a
 * `data-teamedit-image-ratio` attribute.  Safe to call repeatedly: we always
 * recompute from the image's natural dimensions.
 */
export function applyImageRatios(
  root: Document | HTMLElement,
  options: ApplyImageRatiosOptions = {},
) {
  const scope = root instanceof Document ? root : root;
  const images = scope.querySelectorAll<HTMLImageElement>(`img[${IMAGE_RATIO_ATTR}]`);
  for (const image of images) {
    trackImage(image, options);
  }
}
