# 1、研究背景与研究目标

本示例是 md2docx 的测试样例与效果演示，内容为虚构，仅用于覆盖各类转换陷阱。
每个小节末尾用注释标注了它专门用来触发哪个问题，改动本文件时请保留这些结构。

## 1.1 研究目标

本项目面向边缘计算环境下的任务调度问题，研究有限计算资源约束下的任务分配与
性能保障方法。记任务集合为 $\mathcal{K}=\{1,\dots,K\}$，任务 $k$ 的到达速率为
$\lambda_k$、服务速率为 $\mu_k$，系统总计算能力为 $C_{\rm total}$。

其中 $C_{\rm total}$ 与 $\rho_{\rm sys}$ 采用老式字体命令书写，用于验证
`\rm` 是否被正确转换为 `\mathrm`。

<!-- 陷阱 1：\rm 等老式字体命令。texmath 不支持，遇到会把整块公式退化成 LaTeX 原文 -->

## 1.2 系统模型

考虑单节点排队模型，系统利用率定义为

$$
\rho_{\rm sys}
=
\sum_{k\in\mathcal K}
\frac{\lambda_k}{\mu_k}
\tag{1}
$$

稳定性条件要求 $\rho_{\rm sys}<1$。由式（1）可知，当任务数量增加时系统利用率
单调上升。

平均响应时间由排队时延与服务时延两部分构成：

$$
T_k
=
\frac{1}{\mu_k-\lambda_k}
+
d_k^{\rm net}
\tag{2}
$$

其中 $d_k^{\rm net}$ 为网络传输时延。

---

# 2、算法设计与求解方法

## 2.1 任务分配优化

### 2.1.1 优化问题构建

以加权平均响应时间最小为目标，构建如下优化问题。目标函数与约束条件采用
`aligned` 环境书写多行对齐：

$$
\begin{aligned}
\min_{\mathbf x}
\quad&
\sum_{k\in\mathcal K}
\omega_k T_k(x_k)\\
\mathrm{s.t.}
\quad&
\sum_{k\in\mathcal K} x_k \leq C_{\rm total},\\
&
x_k \geq \lambda_k/\mu_k,
\quad
\forall k\in\mathcal K
\end{aligned}
\tag{3}
$$

式（3）中 $\omega_k$ 为任务权重，$x_k$ 为分配给任务 $k$ 的计算资源。

<!-- 陷阱 2：aligned 的对齐符 &。pandoc 会把它当普通字符写进 OMML，公式里出现可见的 & -->

### 2.1.2 分段决策规则

资源分配采用阈值型规则，用 `cases` 环境表示：

$$
x_k^\star
=
\begin{cases}
\lambda_k/\mu_k, & \omega_k \leq \vartheta,\\
C_{\rm total}/K, & \omega_k > \vartheta
\end{cases}
\tag{4}
$$

由式（4）可见，权重超过阈值 $\vartheta$ 的任务获得均分资源。

## （1）性能边界分析

本节标题在源文件里故意写成二级标题，挂在三级标题之下，用于验证标题层级归一化。

<!-- 陷阱 3：标题层级倒挂。### 下面挂 ##，转 Word 后大纲和自动编号会乱 -->

综合上述结果，可得加权响应时间的上下界夹逼关系：

$$
\boxed{
T_{\rm lb}
\leq
T^\star
\leq
T_{\rm ub}
}
\tag{5}
$$

<!-- 陷阱 4：\boxed。OMML 没有对应结构 -->

下面这个公式块内部**故意留了一个空行**。markdown 里空行会断段，pandoc 因此拿不到
闭合的 `$$`，整块公式会被静默吞掉且不报任何警告：

$$
\begin{aligned}
T_{\rm ub}
=&
\max_{k\in\mathcal K}
\frac{1}{\mu_k-\lambda_k}

\\
&+
\max_{k\in\mathcal K}
d_k^{\rm net}
\end{aligned}
\tag{6}
$$

式（6）给出了响应时间上界的显式表达。

<!-- 陷阱 5：公式块内部的空行。这是最危险的一个，会整块丢公式且完全静默 -->

---

## 2.2 求解方法

> **下一步，将基于上述思路重点开展以下研究：**

首先研究凸松弛条件。当权重满足 $\omega_k>0$ 且服务速率满足 $\mu_k>\lambda_k$ 时，
问题（3）为凸问题，可用内点法求解。

下面这段把正文和行间公式写在同一段落（中间没有空行），其中 $$\nabla f(\mathbf x)=\mathbf 0$$ 为一阶最优条件，而 $$\nabla^2 f(\mathbf x)\succ 0$$ 为二阶充分条件，两个行间公式出现在同一段落里。

<!-- 陷阱 6：正文与行间公式混排，且同一段落里有两个行间公式 -->

对于非凸情形，采用如下迭代更新，其中花体、黑板体、哥特体符号并存：

$$
\mathbf x^{(t+1)}
=
\Pi_{\mathcal X}
\left(
\mathbf x^{(t)}
-
\eta_t
\nabla f(\mathbf x^{(t)})
\right),
\qquad
\mathcal X\subseteq\mathbb R^{K},
\quad
\alpha\in\mathfrak A
\tag{7}
$$

收敛性分析中用到的性能等级序列包含中文，用于验证 `\text{}` 里的中文不被误算：

$$
\text{额定保障}
\rightarrow
\text{弹性保障}
\rightarrow
\text{选择性丢弃}
\tag{8}
$$

## 2.3 实验验证

下图为不同任务数下的平均响应时间对比。这里用的是 Obsidian 的 wikilink 图片语法，
且文件故意不存在，用于验证缺图占位提示：

![[performance_comparison.png]]

**图 1：不同任务数下的平均响应时间对比**

下面是一个手写的图占位标记：

【图占位1：系统架构与模块交互示意图】

**表 1：仿真参数设置**

| 参数 | 取值 |
|---|---|
| 任务数 $K$ | 10 |
| 总算力 $C_{\rm total}$ | 100 |
| 权重范围 $\omega_k$ | [0.1, 1.0] |

正文里还有一处引用指向不存在的编号：式（99）——它对不上任何 `\tag`，
预处理会原样保留并报警，不会硬猜。

<!-- 陷阱 7：Obsidian ![[]] 图片语法、图占位、悬空交叉引用 -->
<!-- 陷阱 8：文中多处 --- 分隔线。pandoc 的 multiline_tables 会把它当表格头分隔符，
     从而把后面整段正文吞成表格里的纯文本，连 ## 都不解析 -->
