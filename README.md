# 时代蜂族培训数据追踪

## 当前架构
- 前端页面：`dashboard.html` + `dashboard.js`
- 后端接口：Cloudflare Pages Functions，统一走同域 `/api/*`
- 数据源：Bmob
- 前端不再直连 Bmob，也不再暴露 Bmob Key

## 当前已实现
- 教务入口登录
- 登录后进入教务后台
- 后台可执行：
  - 上传培训数据
  - 导出培训报表
  - 新增管理员账号
  - 页面 UI 缩放比例拖动与锁定
  - 数据变动日志查看
  - 删除已导入文件
  - 删除管理员账号
- 上传 Excel 后会逐行读取各单元格数据，并通过 `/api/upload` 写入 Bmob `TrainingRecords`
- 页面加载时通过 `/api/state` 拉取云端数据并渲染图表

## 当前分类规则
- 每条数据按表头 `课堂类别` 的值分类：
  - `线下课堂` => 归类到线下课堂
  - `自主培训课程` => 归类到自主培训课程
  - `商学院` => 归类到商学院
  - `通关` => 归类到通关
  - 其他值 => 归类到未分类

## 当前时间提取规则
系统会优先从以下表头对应的列里提取时间：
- 时间
- 日期
- 培训时间
- 上课时间
- 开始时间
- 开课时间
- 创建时间
- 提交时间

如果上述表头未命中，会继续扫描整行单元格文本中的日期时间内容。

## 目录说明
- `dashboard.html`：页面结构
- `dashboard.js`：前端业务逻辑，调用 `/api/*`
- `functions/_lib/api.js`：Cloudflare Functions 公共工具、Bmob 调用、鉴权
- `functions/api/login.js`：登录接口
- `functions/api/state.js`：页面状态加载接口
- `functions/api/upload.js`：上传培训数据接口
- `functions/api/files/[id].js`：删除文件接口
- `functions/api/admins.js`：新增管理员接口
- `functions/api/admins/[id].js`：删除管理员接口
- `functions/api/logs.js`：日志写入与清空接口
- `.dev.vars.example`：本地调试所需环境变量模板

## Cloudflare 环境变量
部署到 Cloudflare Pages / Functions 时，至少需要配置：
- `BMOB_APP_ID`
- `BMOB_REST_API_KEY`
- `AUTH_SECRET`
- `SUPER_ADMIN_USERNAME`
- `SUPER_ADMIN_PASSWORD`

其中：
- `AUTH_SECRET` 用来签发登录令牌
- `SUPER_ADMIN_USERNAME` / `SUPER_ADMIN_PASSWORD` 用来在 `Admins` 表为空时自动补一个超级管理员种子账号

## 上传到 GitHub 的文件
需要上传这些内容：
- `dashboard.html`
- `dashboard.js`
- `functions/`
- `README.md`
- `.gitignore`
- `.dev.vars.example`

不要上传这些内容：
- `.dev.vars`
- 任何真实密钥、真实管理员密码

## 部署说明
1. 把上述文件推送到 GitHub 仓库
2. 在 Cloudflare Pages 里连接该 GitHub 仓库
3. 构建命令留空
4. 输出目录设为仓库根目录
5. 在 Cloudflare Pages 的环境变量里填入上面的 5 个变量
6. 部署完成后，前端页面会直接请求同域 `/api/*`

## 本地调试说明
- 如果只是用 `file:///.../dashboard.html` 或普通静态服务器直接打开页面，而没有一起运行 Cloudflare Functions，本地登录仍会失败
- 本地要调接口，需使用 Cloudflare Pages Functions 的本地开发方式运行
- `.dev.vars.example` 复制成 `.dev.vars` 后，再启动 Cloudflare 本地开发服务
