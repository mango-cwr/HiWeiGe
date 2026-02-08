# 将本项目部署到 Vercel

本项目已适配 Vercel 的 Flask 运行方式，上传文件在请求内使用 `/tmp` 目录，无需额外配置即可部署。

## 方式一：通过 Vercel 网页（推荐）

### 1. 准备代码仓库

- 确保当前项目已推送到 GitHub（例如 `https://github.com/mango-cwr/HiWeiGe`）。
- 仓库根目录需包含：`app.py`、`index.html`、`requirements.txt`。

### 2. 在 Vercel 创建项目

1. 打开 [vercel.com](https://vercel.com)，登录（可用 GitHub 账号）。
2. 点击 **Add New…** → **Project**。
3. 在 **Import Git Repository** 中选择你的仓库（如 `mango-cwr/HiWeiGe`），点击 **Import**。

### 3. 配置项目（一般可保持默认）

- **Framework Preset**：选 **Flask**（若没有可选 **Other**）。
- **Root Directory**：留空（使用仓库根目录）。
- **Build Command**：留空或由 Vercel 自动识别。
- **Output Directory**：留空。
- **Install Command**：留空（Vercel 会用 `pip install -r requirements.txt`）。

无需设置环境变量即可运行；若以后需要可在此页添加。

### 4. 部署

点击 **Deploy**，等待构建与部署完成。

### 5. 访问

- 部署成功后，Vercel 会给出一个地址，例如：`https://hi-wei-ge-xxx.vercel.app`。
- 在浏览器打开该地址即可使用「套餐详情分析」和「月度账单差异对比」。

---

## 方式二：通过 Vercel CLI

### 1. 安装 Vercel CLI

```bash
npm i -g vercel
```

### 2. 登录

在项目目录下执行：

```bash
vercel login
```

按提示用浏览器完成登录。

### 3. 部署

在项目根目录（与 `app.py` 同级）执行：

```bash
vercel
```

首次会询问项目名、是否链接到已有 Vercel 项目等，按需选择；之后可直接：

```bash
vercel --prod
```

将当前代码部署到生产环境。

### 4. 本地用 Vercel 环境调试（可选）

```bash
pip install -r requirements.txt
vercel dev
```

在本地用 Vercel 的运行时模拟环境测试。

---

## 注意事项

- **执行时间**：免费版单次请求约 10 秒限制，一般 Excel 解析足够；超大文件可能超时。
- **上传大小**：注意 Vercel 对请求体大小的限制（约 4.5MB），过大的 Excel 可能被拒绝。
- **冷启动**：长时间无访问后第一次请求可能稍慢（加载 Python 与依赖）。
- 代码中已根据 `VERCEL` 环境变量自动使用 `/tmp` 作为上传目录，无需再配置。

---

## 若部署失败可检查

1. 仓库根目录是否有 `app.py`、`requirements.txt`。
2. `app.py` 中 Flask 实例变量名是否为 `app`。
3. 在 Vercel 项目 **Settings → General** 中确认 **Framework Preset** 为 **Flask**（或能识别 Python 的选项）。
