import {
  useEffect,
  useId,
  useLayoutEffect,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent,
} from "react";
import { createPortal } from "react-dom";
import { Check, ChevronDown, ScanLine } from "lucide-react";
import type { WorkspaceOcrModelOption } from "../workspace/workspaceTypes";

interface OcrModelSelectorProps {
  value: string;
  options: readonly WorkspaceOcrModelOption[];
  disabled: boolean;
  isEn: boolean;
  onChange?: (value: string) => void;
}

interface MenuPosition {
  left: number;
  top: number;
  width: number;
}

export function OcrModelSelector({
  value,
  options,
  disabled,
  isEn,
  onChange,
}: OcrModelSelectorProps) {
  const [open, setOpen] = useState(false);
  const [activeIndex, setActiveIndex] = useState(0);
  const [menuPosition, setMenuPosition] = useState<MenuPosition | null>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const listboxId = useId();
  const selectedIndex = Math.max(
    0,
    options.findIndex((option) => option.id === value),
  );
  const selected = options[selectedIndex] ?? options[0];
  const selectedLabel = selected
    ? isEn
      ? selected.selectedLabelEn ?? selected.labelEn
      : selected.selectedLabelVi ?? selected.labelVi
    : "";

  const close = (restoreFocus = false) => {
    setOpen(false);
    setMenuPosition(null);
    if (restoreFocus) triggerRef.current?.focus();
  };

  const openMenu = (index = selectedIndex) => {
    if (disabled || options.length === 0) return;
    setActiveIndex(Math.max(0, Math.min(options.length - 1, index)));
    setOpen(true);
  };

  const selectIndex = (index: number) => {
    const option = options[index];
    if (!option) return;
    onChange?.(option.id);
    close(true);
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLButtonElement>) => {
    if (disabled || options.length === 0) return;
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      const direction = event.key === "ArrowDown" ? 1 : -1;
      if (!open) {
        openMenu(selectedIndex);
      } else {
        setActiveIndex((current) =>
          (current + direction + options.length) % options.length,
        );
      }
      return;
    }
    if (event.key === "Home" && open) {
      event.preventDefault();
      setActiveIndex(0);
      return;
    }
    if (event.key === "End" && open) {
      event.preventDefault();
      setActiveIndex(options.length - 1);
      return;
    }
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      if (open) selectIndex(activeIndex);
      else openMenu();
      return;
    }
    if (event.key === "Escape" && open) {
      event.preventDefault();
      close(true);
      return;
    }
    if (event.key === "Tab" && open) {
      close();
    }
  };

  useEffect(() => {
    if (!open) return;
    const handlePointerDown = (event: PointerEvent) => {
      if (!(event.target instanceof Node)) return;
      if (
        triggerRef.current?.contains(event.target) ||
        menuRef.current?.contains(event.target)
      ) {
        return;
      }
      close();
    };
    document.addEventListener("pointerdown", handlePointerDown, true);
    return () => document.removeEventListener("pointerdown", handlePointerDown, true);
  }, [open]);

  useEffect(() => {
    if (disabled && open) close();
  }, [disabled, open]);

  useLayoutEffect(() => {
    if (!open) return;
    let frame = 0;
    const updatePosition = () => {
      window.cancelAnimationFrame(frame);
      frame = window.requestAnimationFrame(() => {
        const trigger = triggerRef.current;
        if (!trigger) return;
        const rect = trigger.getBoundingClientRect();
        const margin = 8;
        const gap = 6;
        const width = Math.min(
          240,
          Math.max(rect.width, Math.min(210, window.innerWidth - margin * 2)),
        );
        const left = Math.min(
          Math.max(margin, rect.right - width),
          Math.max(margin, window.innerWidth - width - margin),
        );
        const estimatedHeight = Math.min(180, options.length * 38 + 10);
        const below = rect.bottom + gap;
        const top =
          below + estimatedHeight <= window.innerHeight - margin
            ? below
            : Math.max(margin, rect.top - estimatedHeight - gap);
        setMenuPosition({ left, top, width });
      });
    };
    window.addEventListener("resize", updatePosition);
    window.addEventListener("scroll", updatePosition, true);
    updatePosition();
    return () => {
      window.cancelAnimationFrame(frame);
      window.removeEventListener("resize", updatePosition);
      window.removeEventListener("scroll", updatePosition, true);
    };
  }, [open, options.length]);

  const menuStyle = menuPosition
    ? ({
        left: `${menuPosition.left}px`,
        top: `${menuPosition.top}px`,
        width: `${menuPosition.width}px`,
      } as CSSProperties)
    : undefined;

  return (
    <div
      className="canvas-ocr-model"
      title={
        isEn
          ? "Model used when an image is pasted into a formula field"
          : "Model được sử dụng khi dán hình ảnh vào trường công thức"
      }
    >
      <ScanLine size={14} />
      <button
        ref={triggerRef}
        type="button"
        className="ocr-model-selector-trigger"
        disabled={disabled}
        role="combobox"
        aria-expanded={open}
        aria-controls={listboxId}
        aria-haspopup="listbox"
        aria-activedescendant={open ? `${listboxId}-option-${activeIndex}` : undefined}
        aria-label={isEn ? "OCR recognition model" : "Mô hình nhận dạng OCR"}
        onClick={() => (open ? close() : openMenu())}
        onKeyDown={handleKeyDown}
        data-ocr-model-trigger
      >
        <span>{selectedLabel}</span>
        <ChevronDown size={12} aria-hidden="true" />
      </button>
      {open &&
        menuPosition &&
        createPortal(
          <div
            ref={menuRef}
            id={listboxId}
            className="ocr-model-selector-menu"
            role="listbox"
            aria-label={isEn ? "OCR recognition model" : "Mô hình nhận dạng OCR"}
            style={menuStyle}
            data-ocr-model-menu
          >
            {options.map((option, index) => {
              const selectedOption = option.id === value;
              return (
                <button
                  id={`${listboxId}-option-${index}`}
                  key={option.id}
                  type="button"
                  className={index === activeIndex ? "is-active" : ""}
                  role="option"
                  aria-selected={selectedOption}
                  onPointerEnter={() => setActiveIndex(index)}
                  onClick={() => selectIndex(index)}
                  data-ocr-model-option={option.id}
                >
                  <span>{isEn ? option.labelEn : option.labelVi}</span>
                  {selectedOption && <Check size={14} aria-hidden="true" />}
                </button>
              );
            })}
          </div>,
          document.body,
        )}
    </div>
  );
}
