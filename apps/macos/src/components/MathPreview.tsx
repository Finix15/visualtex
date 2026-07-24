import { useLayoutEffect, useMemo, useRef } from "react";
import { convertLatexToMarkup } from "mathlive";

interface MathPreviewProps {
  latex: string;
  className?: string;
  fit?: boolean;
}

const fitInsetRatio = 0.9;
const minimumFitScale = 0.1;
const maximumFitScale = 8;

export function MathPreview({
  latex,
  className = "",
  fit = false,
}: MathPreviewProps) {
  const hostRef = useRef<HTMLSpanElement>(null);
  const contentRef = useRef<HTMLSpanElement>(null);
  const markup = useMemo(
    () => convertLatexToMarkup(latex, { defaultMode: "math" }),
    [latex],
  );

  useLayoutEffect(() => {
    const host = hostRef.current;
    const content = contentRef.current;
    if (!host || !content) return;

    let animationFrame = 0;
    const measure = () => {
      animationFrame = 0;
      if (!fit) {
        content.style.setProperty("--math-preview-fit-scale", "1");
        host.dataset.fitReady = "false";
        host.dataset.fitScale = "1";
        return;
      }

      const availableWidth = Math.max(1, host.clientWidth * fitInsetRatio);
      const availableHeight = Math.max(1, host.clientHeight * fitInsetRatio);
      const naturalWidth = Math.max(1, content.offsetWidth);
      const naturalHeight = Math.max(1, content.offsetHeight);
      const containedScale = Math.min(
        availableWidth / naturalWidth,
        availableHeight / naturalHeight,
      );
      const scale = Math.max(
        minimumFitScale,
        Math.min(maximumFitScale, containedScale),
      );

      content.style.setProperty(
        "--math-preview-fit-scale",
        scale.toFixed(4),
      );
      host.dataset.fitReady = "true";
      host.dataset.fitScale = scale.toFixed(4);
    };
    const scheduleMeasure = () => {
      if (animationFrame) cancelAnimationFrame(animationFrame);
      animationFrame = requestAnimationFrame(measure);
    };

    scheduleMeasure();
    void document.fonts?.ready.then(scheduleMeasure);
    const resizeObserver = new ResizeObserver(scheduleMeasure);
    resizeObserver.observe(host);

    return () => {
      if (animationFrame) cancelAnimationFrame(animationFrame);
      resizeObserver.disconnect();
    };
  }, [fit, markup]);

  return (
    <span
      ref={hostRef}
      className={"math-preview " + className}
      aria-hidden="true"
      data-fit={fit ? "contain" : "none"}
    >
      <span
        ref={contentRef}
        className="math-preview-fit-content"
        dangerouslySetInnerHTML={{ __html: markup }}
      />
    </span>
  );
}
