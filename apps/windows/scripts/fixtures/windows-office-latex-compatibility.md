# VisualTeX LaTeX 兼容与尺寸验收

正文 11 磅基准字符 x，行内简单公式 $x$ 后文字保持原生；普通分式 $\frac{a+b}{c+d}$ 与嵌套上下标 $T_{i_{\text{内层}}}^{n^2}$ 均需可编辑。

行内粗斜体向量 $A\bm v=\lambda\bm v$，以及 $\boldsymbol{\alpha}+\mathbf{x}+\mathrm{d}+\operatorname{rank}(A)$ 不得出现原始命令文本。

行内数学字体 $\mathbb{R}+\mathcal{L}+\mathfrak{g}$，重音 $\vec{x}+\hat{y}+\bar{z}+\dot{q}+\ddot{r}$；行内裸积分 $\int f(x)\,\mathrm{d}x$ 不得显示空上下限占位框。

$$
\nabla\times\bm F=\begin{vmatrix}
\bm e_x&\bm e_y&\bm e_z\\
\partial_x&\partial_y&\partial_z\\
F_x&F_y&F_z
\end{vmatrix}.
$$

$$
f(x)=\begin{cases}
\sqrt{x^2+1},&x\ge 0,\\
-\left(\dfrac{1+x}{1-x}\right),&x<0.
\end{cases}
$$

$$
\sum_{\substack{i=1\\j=1}}^{n}a_{ij}+\int_{0}^{1}\frac{x^2}{\sqrt{1-x^2}}\,\mathrm{d}x
$$

$$
\nabla f=\left(\frac{\partial f}{\partial x},\frac{\partial f}{\partial y},\frac{\partial f}{\partial z}\right),\qquad
\nabla\cdot\bm F=\frac{\partial F_x}{\partial x}+\frac{\partial F_y}{\partial y}+\frac{\partial F_z}{\partial z}.
$$

$$
\nabla^2 f=\frac{\partial^2 f}{\partial x^2}+\frac{\partial^2 f}{\partial y^2}+\frac{\partial^2 f}{\partial z^2}.
$$

$$
\iiint_V\nabla\cdot\bm F\,\mathrm{d}V=\iint_{\partial V}\bm F\cdot\mathrm{d}\bm S.
$$

$$
\iint_S(\nabla\times\bm F)\cdot\mathrm{d}\bm S=\oint_{\partial S}\bm F\cdot\mathrm{d}\bm l.
$$

$$
\begin{aligned}
\alpha+\beta+\gamma&=\pi,\\
\overset{\text{上}}{X}+\underset{\text{下}}{Y}&=\left\lVert\bm v\right\rVert.
\end{aligned}
$$
