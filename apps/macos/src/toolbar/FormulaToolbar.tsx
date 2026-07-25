import { useEffect, useMemo, useState, type MouseEvent } from "react";
import { Star } from "lucide-react";
import type { LatexCommand } from "../types/command";
import {
  categoryLabels,
  categoryLabelsEn,
  calculusCommandIds,
  commandRegistry,
  commonCommandIds,
} from "../autocomplete/commandRegistry";
import { MathPreview } from "../components/MathPreview";
import { useEditorStore } from "../stores/editorStore";

interface Props {
  onInsert: (command: LatexCommand) => void;
}

type MatrixDelimiter = "bmatrix" | "pmatrix" | "vmatrix";
type ToolbarView = "tools" | "tiles";
type TileCategory = "custom" | "common";

interface FormulaTileDefinition {
  id: string;
  latex: string;
  labelZh: string;
  labelEn: string;
}

interface FormulaTileContextMenuState {
  latex: string;
  x: number;
  y: number;
}

const customFormulaTilesStorageKey = "visualtex-custom-formula-tiles";

const commonFormulaTiles: FormulaTileDefinition[] = [
  {
    id: "quadratic-formula",
    latex: "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}",
    labelZh: "一元二次方程求根公式",
    labelEn: "Quadratic formula",
  },
  {
    id: "euler-identity",
    latex: "e^{i\\pi}+1=0",
    labelZh: "欧拉恒等式",
    labelEn: "Euler identity",
  },
  {
    id: "pythagorean-theorem",
    latex: "a^2+b^2=c^2",
    labelZh: "勾股定理",
    labelEn: "Pythagorean theorem",
  },
  {
    id: "binomial-theorem",
    latex: "(a+b)^n=\\sum_{k=0}^{n}\\binom{n}{k}a^{n-k}b^k",
    labelZh: "二项式定理",
    labelEn: "Binomial theorem",
  },
  {
    id: "gaussian-integral",
    latex: "\\int_{-\\infty}^{\\infty}e^{-x^2}\\,\\mathrm{d}x=\\sqrt{\\pi}",
    labelZh: "高斯积分",
    labelEn: "Gaussian integral",
  },
  {
    id: "taylor-series",
    latex: "f(x)=\\sum_{n=0}^{\\infty}\\frac{f^{(n)}(a)}{n!}(x-a)^n",
    labelZh: "泰勒展开",
    labelEn: "Taylor series",
  },
  {
    id: "mass-energy",
    latex: "E=mc^2",
    labelZh: "质能方程",
    labelEn: "Mass-energy equivalence",
  },
  {
    id: "schrodinger-equation",
    latex: "i\\hbar\\frac{\\partial}{\\partial t}\\Psi=\\hat{H}\\Psi",
    labelZh: "含时薛定谔方程",
    labelEn: "Time-dependent Schrödinger equation",
  },
  {
    id: "gauss-law",
    latex: "\\nabla\\cdot\\mathbf{E}=\\frac{\\rho}{\\varepsilon_0}",
    labelZh: "高斯定律",
    labelEn: "Gauss's law",
  },
  {
    id: "characteristic-equation",
    latex: "\\det(A-\\lambda I)=0",
    labelZh: "矩阵特征方程",
    labelEn: "Matrix characteristic equation",
  },
];

function loadCustomFormulaTiles() {
  try {
    const stored = JSON.parse(
      localStorage.getItem(customFormulaTilesStorageKey) ?? "[]",
    );
    if (!Array.isArray(stored)) return [];
    return stored
      .filter((value): value is string => typeof value === "string")
      .map((value) => value.trim())
      .filter(Boolean)
      .slice(0, 30);
  } catch {
    return [];
  }
}

function persistCustomFormulaTiles(tiles: readonly string[]) {
  try {
    localStorage.setItem(customFormulaTilesStorageKey, JSON.stringify(tiles));
  } catch {
    // Keep the current session usable even when storage is unavailable.
  }
}

const categories = [
  "common",
  "structure",
  "calculus",
  "matrix",
  "relation",
  "greek",
  "arrow",
  "physics",
  "set",
];

const matrixGridCells = Array.from({ length: 100 }, (_, index) => ({
  row: Math.floor(index / 10) + 1,
  column: (index % 10) + 1,
}));
const matrixDelimiterOptions: Array<{
  id: MatrixDelimiter;
  preview: string;
  labelZh: string;
  labelEn: string;
}> = [
  {
    id: "vmatrix",
    preview: "\\begin{vmatrix}a&b\\\\c&d\\end{vmatrix}",
    labelZh: "竖线",
    labelEn: "Bars",
  },
  {
    id: "bmatrix",
    preview: "\\begin{bmatrix}a&b\\\\c&d\\end{bmatrix}",
    labelZh: "方括号",
    labelEn: "Brackets",
  },
  {
    id: "pmatrix",
    preview: "\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}",
    labelZh: "圆括号",
    labelEn: "Parentheses",
  },
];

const autoFitCategories = new Set([
  "common",
  "structure",
  "calculus",
  "matrix",
]);
const physicsAutoFitCommandIds = new Set([
  "commutator",
  "anticommutator",
]);

const calculusPreviewById: Record<string, string> = {
  intplain: "\\int",
  int: "\\int_a^b",
  "iint-bounds": "\\iint_D^S",
  "iiint-bounds": "\\iiint_V^W",
  "oint-bounds": "\\oint_C^D",
  lineintegral: "\\int_C",
  iint: "\\iint_D",
  surfaceintegral: "\\iint_S",
  iiint: "\\iiint_V",
  volumeintegral: "\\iiint_V",
  oint: "\\oint_C",
  "closed-surface-integral": "\\oiint_S",
  "closed-volume-integral": "\\oiiint_V",
  sum: "\\sum_{i=1}^{n}",
  "sum-finite": "\\sum_{k=1}^{n}",
  series: "\\sum_{n=0}^{\\infty}",
  prod: "\\prod_{i=1}^{n}",
  "prod-finite": "\\prod_{k=1}^{n}",
  productseries: "\\prod_{n=1}^{\\infty}",
  coproduct: "\\coprod_{i=1}^{n}",
  lim: "\\lim_{x\\to0}",
  "lim-infty": "\\lim_{x\\to\\infty}",
  "lim-left": "\\lim_{x\\to a^-}",
  "lim-right": "\\lim_{x\\to a^+}",
  derivative: "\\frac{\\mathrm{d}}{\\mathrm{d}x}",
  secondderivative: "\\frac{\\mathrm{d}^{2}}{\\mathrm{d}x^{2}}",
  partial: "\\frac{\\partial}{\\partial x}",
  partialsecond: "\\frac{\\partial^{2}}{\\partial x^{2}}",
  mixedpartial: "\\frac{\\partial^{2}}{\\partial x\\partial y}",
  evalbar: "\\left.\\vphantom{F}\\right|_a^b",
  nabla: "\\nabla",
  ln: "\\ln",
  log: "\\log_a",
  exp: "\\exp",
  sin: "\\sin",
  cos: "\\cos",
  tan: "\\tan",
};

const toolbarPreviewById: Record<string, string> = {
  ...calculusPreviewById,
  cases: "\\begin{cases}a\\\\b\\end{cases}",
  overbrace: "\\overbrace{a+b}",
  underbrace: "\\underbrace{a+b}",
  rowvector: "\\begin{bmatrix}a&b\\end{bmatrix}",
  colvector: "\\begin{bmatrix}a\\\\b\\end{bmatrix}",
  det: "\\det",
  trace: "\\operatorname{tr}",
  rank: "\\operatorname{rank}",
  transpose: "A^{\\mathsf{T}}",
  inverse: "A^{-1}",
  dotproduct: "\\bullet",
};

const complexPreviewTokens = [
  "\\sum",
  "\\prod",
  "\\coprod",
  "\\int",
  "\\iint",
  "\\iiint",
  "\\oint",
  "\\oiint",
  "\\oiiint",
  "\\lim",
];
const mediumPreviewTokens = [
  "\\frac",
  "\\dfrac",
  "\\tfrac",
  "\\sqrt",
  "\\binom",
  "\\left",
  "\\right",
];

const previewSizeClass = (command: LatexCommand) => {
  const latex = command.previewLatex;
  const isComplex =
    latex.includes("\\begin") ||
    latex.includes("cases") ||
    complexPreviewTokens.some((token) => latex.includes(token)) ||
    command.id === "derivative" ||
    command.id === "partial";

  if (isComplex) return " is-complex";

  const visibleAtomCount = latex
    .replace(/\\[A-Za-z]+/g, "x")
    .replace(/[{}_^()[\]\\|\s]/g, "").length;
  const isMedium =
    mediumPreviewTokens.some((token) => latex.includes(token)) ||
    visibleAtomCount > 3;

  return isMedium ? " is-medium" : " is-symbol-large";
};

const toolbarPreviewLatex = (command: LatexCommand) =>
  toolbarPreviewById[command.id] ?? command.previewLatex;

const shouldAutoFitPreview = (category: string, commandId: string) =>
  autoFitCategories.has(category) ||
  (category === "physics" && physicsAutoFitCommandIds.has(commandId));

function createTileCommand(tile: FormulaTileDefinition): LatexCommand {
  return {
    id: `formula-tile-${tile.id}`,
    command: tile.latex,
    insertTemplate: tile.latex,
    previewLatex: tile.latex,
    labelZh: tile.labelZh,
    labelEn: tile.labelEn,
    aliases: ["tile", "formula"],
    keywords: ["磁贴", "公式"],
    category: "structure",
    defaultPriority: 120,
    supportedInMathMode: true,
  };
}

function createMatrixCommand(
  rows: number,
  columns: number,
  delimiter: MatrixDelimiter,
): LatexCommand {
  const matrixBody = Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => "\\placeholder{}").join(" & "),
  ).join(" \\\\ ");
  const delimiterCopy = matrixDelimiterOptions.find(
    (option) => option.id === delimiter,
  ) ?? matrixDelimiterOptions[1];

  return {
    id: `custom-${delimiter}-${rows}x${columns}`,
    command: `\\begin{${delimiter}}`,
    insertTemplate: `\\begin{${delimiter}}${matrixBody}\\end{${delimiter}}`,
    previewLatex: delimiterCopy.preview,
    labelZh: `${rows}×${columns} ${delimiterCopy.labelZh}矩阵`,
    labelEn: `${rows}×${columns} ${delimiterCopy.labelEn.toLowerCase()} matrix`,
    aliases: ["matrix", delimiter],
    keywords: ["矩阵", "自定义矩阵", `${rows}x${columns}`],
    category: "matrix",
    defaultPriority: 120,
    supportedInMathMode: true,
  };
}

export function FormulaToolbar({ onInsert }: Props) {
  const [activeView, setActiveView] = useState<ToolbarView>("tools");
  const [activeTileCategory, setActiveTileCategory] =
    useState<TileCategory>("common");
  const [customFormulaTiles, setCustomFormulaTiles] = useState<string[]>(
    loadCustomFormulaTiles,
  );
  const [tileContextMenu, setTileContextMenu] =
    useState<FormulaTileContextMenuState | null>(null);
  const [activeCategory, setActiveCategory] = useState("common");
  const [matrixRows, setMatrixRows] = useState(2);
  const [matrixColumns, setMatrixColumns] = useState(2);
  const [matrixHover, setMatrixHover] = useState<{
    rows: number;
    columns: number;
  } | null>(null);
  const [matrixDelimiter, setMatrixDelimiter] =
    useState<MatrixDelimiter>("bmatrix");
  const language = useEditorStore((state) => state.language);
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const isEn = language === "en";
  const activeLineLatex = useMemo(
    () => lines.find((line) => line.id === activeLineId)?.latex.trim() ?? "",
    [activeLineId, lines],
  );

  useEffect(() => {
    if (!tileContextMenu) return;

    const closeFromPointer = (event: PointerEvent) => {
      const target = event.target;
      if (
        target instanceof Element &&
        target.closest(".formula-tile-context-menu")
      ) {
        return;
      }
      setTileContextMenu(null);
    };
    const closeFromKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") setTileContextMenu(null);
    };
    const closeMenu = () => setTileContextMenu(null);

    document.addEventListener("pointerdown", closeFromPointer);
    document.addEventListener("keydown", closeFromKey);
    window.addEventListener("blur", closeMenu);
    window.addEventListener("resize", closeMenu);
    window.addEventListener("scroll", closeMenu, true);

    return () => {
      document.removeEventListener("pointerdown", closeFromPointer);
      document.removeEventListener("keydown", closeFromKey);
      window.removeEventListener("blur", closeMenu);
      window.removeEventListener("resize", closeMenu);
      window.removeEventListener("scroll", closeMenu, true);
    };
  }, [tileContextMenu]);

  const visibleCommands = useMemo(() => {
    const preferredIds = activeCategory === "common"
      ? commonCommandIds
      : activeCategory === "calculus"
        ? calculusCommandIds
        : null;
    if (preferredIds) {
      return preferredIds
        .map((id) => commandRegistry.find((command) => command.id === id))
        .filter((command): command is LatexCommand => Boolean(command));
    }
    return commandRegistry.filter(
      (command) => command.category === activeCategory,
    );
  }, [activeCategory]);

  const customTileDefinitions = useMemo<FormulaTileDefinition[]>(
    () =>
      customFormulaTiles.map((latex, index) => ({
        id: `custom-${index}`,
        latex,
        labelZh: `自定义公式 ${index + 1}`,
        labelEn: `Custom formula ${index + 1}`,
      })),
    [customFormulaTiles],
  );
  const visibleFormulaTiles =
    activeTileCategory === "common"
      ? commonFormulaTiles
      : customTileDefinitions;

  const insertCustomMatrix = () => {
    onInsert(createMatrixCommand(matrixRows, matrixColumns, matrixDelimiter));
  };
  const insertFormulaTile = (tile: FormulaTileDefinition) => {
    onInsert(createTileCommand(tile));
  };
  const saveActiveFormulaAsTile = () => {
    if (!activeLineLatex) return;
    setCustomFormulaTiles((current) => {
      const next = [
        activeLineLatex,
        ...current.filter((latex) => latex !== activeLineLatex),
      ].slice(0, 30);
      persistCustomFormulaTiles(next);
      return next;
    });
    setActiveTileCategory("custom");
  };
  const openCustomTileContextMenu = (
    event: MouseEvent<HTMLButtonElement>,
    tile: FormulaTileDefinition,
  ) => {
    event.preventDefault();
    event.stopPropagation();
    const menuWidth = 152;
    const menuHeight = 42;
    setTileContextMenu({
      latex: tile.latex,
      x: Math.max(8, Math.min(event.clientX, window.innerWidth - menuWidth - 8)),
      y: Math.max(8, Math.min(event.clientY, window.innerHeight - menuHeight - 8)),
    });
  };
  const deleteCustomFormulaTile = (latex: string) => {
    setCustomFormulaTiles((current) => {
      const next = current.filter((item) => item !== latex);
      persistCustomFormulaTiles(next);
      return next;
    });
    setTileContextMenu(null);
  };
  const previewRows = matrixHover?.rows ?? matrixRows;
  const previewColumns = matrixHover?.columns ?? matrixColumns;

  return (
    <aside
      className="formula-toolbar"
      aria-label={isEn ? "Formula toolbar" : "公式工具栏"}
    >
      <header className="formula-toolbar-header">
        <div
          className="formula-toolbar-view-tabs"
          role="tablist"
          aria-label={isEn ? "Sidebar view" : "侧栏视图"}
        >
          <button
            type="button"
            role="tab"
            className={activeView === "tools" ? "is-active" : ""}
            aria-selected={activeView === "tools"}
            data-toolbar-view="tools"
            onClick={() => {
              setTileContextMenu(null);
              setActiveView("tools");
            }}
          >
            {isEn ? "Formula tools" : "公式工具"}
          </button>
          <button
            type="button"
            role="tab"
            className={activeView === "tiles" ? "is-active" : ""}
            aria-selected={activeView === "tiles"}
            data-toolbar-view="tiles"
            onClick={() => {
              setTileContextMenu(null);
              setActiveView("tiles");
            }}
          >
            {isEn ? "Tiles" : "磁贴"}
          </button>
        </div>
      </header>

      {activeView === "tools" ? (
        <>
      <nav className="toolbar-tabs" aria-label={isEn ? "Formula categories" : "公式分类"}>
        {categories.map((category) => (
          <button
            key={category}
            type="button"
            className={
              "toolbar-tab " +
              (activeCategory === category ? "is-active" : "")
            }
            data-category={category}
            aria-pressed={activeCategory === category}
            onClick={() => setActiveCategory(category)}
          >
            {category === "common" && <Star size={13} />}
            {(isEn ? categoryLabelsEn : categoryLabels)[category]}
          </button>
        ))}
      </nav>

      <div className="template-strip" aria-label={isEn ? "Formula templates" : "公式模板"}>
        {activeCategory === "matrix" && (
          <section className="matrix-builder" aria-label={isEn ? "Custom matrix" : "自定义矩阵"}>
            <div className="matrix-builder-heading">
              <strong>{isEn ? "Custom matrix" : "自定义矩阵"}</strong>
              <span className="matrix-size-badge" aria-live="polite">
                {previewRows} × {previewColumns}
              </span>
            </div>

            <div className="matrix-delimiter-options" role="group" aria-label={isEn ? "Matrix delimiter" : "矩阵边界"}>
              {matrixDelimiterOptions.map((option) => (
                <button
                  key={option.id}
                  type="button"
                  className={matrixDelimiter === option.id ? "is-active" : ""}
                  aria-pressed={matrixDelimiter === option.id}
                  onClick={() => setMatrixDelimiter(option.id)}
                  title={isEn ? option.labelEn : option.labelZh}
                  aria-label={isEn ? option.labelEn : option.labelZh}
                >
                  <MathPreview latex={option.preview} fit />
                </button>
              ))}
            </div>

            <div className="matrix-size-picker">
              <div
                className="matrix-size-grid"
                role="grid"
                aria-label={
                  isEn
                    ? "Select matrix rows and columns"
                    : "选择矩阵行数和列数"
                }
                aria-rowcount={10}
                aria-colcount={10}
                onPointerLeave={() => setMatrixHover(null)}
                onBlur={(event) => {
                  if (
                    !event.currentTarget.contains(
                      event.relatedTarget as Node | null,
                    )
                  ) {
                    setMatrixHover(null);
                  }
                }}
              >
                {matrixGridCells.map(({ row, column }) => {
                  const previewed =
                    row <= previewRows && column <= previewColumns;
                  const selectedCorner =
                    row === matrixRows && column === matrixColumns;
                  return (
                    <button
                      key={`${row}-${column}`}
                      type="button"
                      role="gridcell"
                      className={
                        "matrix-size-cell" +
                        (previewed ? " is-previewed" : "") +
                        (selectedCorner ? " is-selected-corner" : "")
                      }
                      aria-label={
                        isEn
                          ? `${row} rows by ${column} columns`
                          : `${row} 行 ${column} 列`
                      }
                      aria-selected={selectedCorner}
                      data-matrix-rows={row}
                      data-matrix-columns={column}
                      onPointerEnter={() =>
                        setMatrixHover({ rows: row, columns: column })
                      }
                      onFocus={() =>
                        setMatrixHover({ rows: row, columns: column })
                      }
                      onClick={() => {
                        setMatrixRows(row);
                        setMatrixColumns(column);
                        setMatrixHover(null);
                      }}
                    />
                  );
                })}
              </div>
            </div>

            <button
              type="button"
              className="matrix-insert-button"
              data-command-id="custom-matrix"
              onClick={insertCustomMatrix}
            >
              {isEn
                ? `Insert ${matrixRows} × ${matrixColumns} matrix`
                : `插入 ${matrixRows} × ${matrixColumns} 矩阵`}
            </button>
          </section>
        )}

        {visibleCommands.map((command) => {
          const autoFit = shouldAutoFitPreview(activeCategory, command.id);
          const previewLatex = toolbarPreviewLatex(command);
          return (
            <button
              type="button"
              className={
                "template-button" +
                (autoFit ? " is-auto-fit" : previewSizeClass(command))
              }
              data-command-id={command.id}
              data-preview-latex={previewLatex}
              key={command.id}
              onClick={() => onInsert(command)}
              aria-label={isEn ? command.labelEn : command.labelZh}
              title={
                (isEn ? command.labelEn : command.labelZh) +
                " · " +
                command.command
              }
            >
              <MathPreview latex={previewLatex} fit={autoFit} />
            </button>
          );
        })}
      </div>
        </>
      ) : (
        <section
          className="formula-tiles-panel"
          aria-label={isEn ? "Formula tiles" : "公式磁贴"}
        >
          <nav
            className="formula-tile-tabs"
            aria-label={isEn ? "Tile categories" : "磁贴分类"}
          >
            <button
              type="button"
              className={activeTileCategory === "custom" ? "is-active" : ""}
              aria-pressed={activeTileCategory === "custom"}
              data-tile-category="custom"
              onClick={() => {
                setTileContextMenu(null);
                setActiveTileCategory("custom");
              }}
            >
              {isEn ? "Custom" : "自定义"}
            </button>
            <button
              type="button"
              className={activeTileCategory === "common" ? "is-active" : ""}
              aria-pressed={activeTileCategory === "common"}
              data-tile-category="common"
              onClick={() => {
                setTileContextMenu(null);
                setActiveTileCategory("common");
              }}
            >
              {isEn ? "Common" : "常用"}
            </button>
          </nav>

          {activeTileCategory === "custom" && (
            <div className="custom-formula-tile-controls">
              <button
                type="button"
                className="save-current-formula-tile"
                disabled={!activeLineLatex}
                onClick={saveActiveFormulaAsTile}
              >
                {isEn
                  ? "Set selected line as custom tile"
                  : "设置为自定义磁贴"}
              </button>
              <span>
                {activeLineLatex
                  ? isEn
                    ? "Uses the currently selected formula line"
                    : "使用当前选中的公式行"
                  : isEn
                    ? "Select a non-empty formula line first"
                    : "请先选择一个非空公式行"}
              </span>
            </div>
          )}

          <div className="formula-tile-list">
            {activeTileCategory === "custom" &&
              visibleFormulaTiles.length === 0 && (
                <div className="formula-tile-empty">
                  <strong>{isEn ? "No custom tiles yet" : "还没有自定义磁贴"}</strong>
                  <span>
                    {isEn
                      ? "Select a formula line, then save it as a tile."
                      : "选中一个公式行，然后点击上方按钮保存。"}
                  </span>
                </div>
              )}
            {visibleFormulaTiles.map((tile) => (
              <button
                type="button"
                className="formula-tile-button"
                key={tile.id}
                data-formula-tile-id={tile.id}
                data-formula-tile-latex={tile.latex}
                onClick={() => insertFormulaTile(tile)}
                onContextMenu={
                  activeTileCategory === "custom"
                    ? (event) => openCustomTileContextMenu(event, tile)
                    : undefined
                }
                aria-label={isEn ? tile.labelEn : tile.labelZh}
                title={
                  activeTileCategory === "custom"
                    ? isEn
                      ? `${tile.labelEn} · Right-click to delete`
                      : `${tile.labelZh} · 右键删除`
                    : isEn
                      ? tile.labelEn
                      : tile.labelZh
                }
              >
                <MathPreview
                  latex={tile.latex}
                  className="formula-tile-preview"
                  fit
                  fluidHeight
                />
              </button>
            ))}
          </div>
        </section>
      )}

      {tileContextMenu && (
        <div
          className="formula-tile-context-menu"
          role="menu"
          aria-label={isEn ? "Custom tile actions" : "自定义磁贴操作"}
          style={{ left: tileContextMenu.x, top: tileContextMenu.y }}
        >
          <button
            type="button"
            role="menuitem"
            onClick={() => deleteCustomFormulaTile(tileContextMenu.latex)}
          >
            {isEn ? "Delete tile" : "删除磁贴"}
          </button>
        </div>
      )}
    </aside>
  );
}
