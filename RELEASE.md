# 版本发布与提交流程（给自己看）

> 这份文档只讲一件事：**以后改了代码，怎么把它传到 GitHub、怎么发新版本。**
> 全程用 GitHub Desktop，不用敲命令。

---

## 一、先搞清两个文件夹的关系（最容易搞混的地方）

| 文件夹 | 是什么 | 谁在用 |
|---|---|---|
| `/Volumes/Port_HDD/.../FILM/000_胶卷索引信息/` | **你平时运行、改代码的地方**，里面有你的 121 卷私人数据 | 你日常用 |
| `~/Documents/film-index/` | **GitHub 仓库文件夹**（此文件夹名仍是旧的 film-index，与仓库名无关） | GitHub Desktop 用 |

> **项目名**：GitHub 上的仓库叫 **`film-index-page`**（页面标题也叫「胶卷索引页面」）。

**改了代码之后，要先把改动同步到仓库文件夹**，GitHub Desktop 才看得到。

→ 双击仓库文件夹里的 **`同步代码.command`** 即可（会自动跳过 `films.xlsx` 等私人文件）。

---

## 二、日常小改动（改几行代码）

```
1. 在硬盘上的软件目录里改代码、跑一遍确认没问题
2. 双击 ~/Documents/film-index/同步代码.command
3. 打开 GitHub Desktop，左边 Changes 会列出改了哪些文件
   （绿色 + 是新增，红色 − 是删除，可以逐个点开看）
4. 左下角 Summary 填一句，比如：
     v0.5.3：修复日历视图横条对齐
   Description 可留空
5. 点 Commit to main
6. 点顶部的 Push origin
```

完成。GitHub 上就多了一条历史记录。

---

## 三、发一个新版本（完整流程）

以 v0.5.3 为例：

### 第 1 步：改版本号（4 个地方）

| 文件 | 改什么 |
|---|---|
| `assets/worker.py` | `VERSION = "0.5.3"` |
| `assets/build_html.py` | 页面角标 `>v0.5.3<` |
| `版本更新.txt` | 在最上面加一条 `- v0.5.3（日期）` + 逐条说明 |
| `README.md` | 「当前版本」那一行 |

改完重新生成页面：`python3 assets/build_html.py`

### 第 2 步：打包 zip

把软件目录打成 `胶卷索引信息_v0.5.3.zip`（**注意：不要包含 films.xlsx / snapshot.json / log / backup**）。

### 第 3 步：同步到仓库 + 提交

```
1. 双击 同步代码.command
2. 顺手把 README.md、「当前版本」也改一下
3. GitHub Desktop → Summary 填：v0.5.3：xxx
4. Commit to main → Push origin
```

### 第 4 步：在 GitHub 上发 Release

```
1. 打开 https://github.com/Anthony-Leo-368/film-index-page
2. 右侧 Releases → Create a new release
3. Choose a tag → 填 v0.5.3 → Create new tag
4. Release title：v0.5.3
5. Describe this release：把 版本更新.txt 里 v0.5.3 那一段粘进来
6. 下面 Attach binaries：把 胶卷索引信息_v0.5.3.zip 拖进去
7. Publish release
```

> Releases 页面会提示「There were no pull requests associated with the commits...」——
> **这是正常的**，因为你是直接提交 main 分支，没有用 PR 流程。忽略它，手写说明即可。

---

## 四、在哪看历史

| 想看什么 | 地址 |
|---|---|
| 所有提交记录 | https://github.com/Anthony-Leo-368/film-index-page/commits/main |
| 所有版本标签 | https://github.com/Anthony-Leo-368/film-index-page/tags |
| 所有发行版（含下载包） | https://github.com/Anthony-Leo-368/film-index-page/releases |
| 某次提交改了什么 | 点进那条提交 → 绿红逐行对比 |
| 本地看 | GitHub Desktop → History |

---

## 五、几条铁律

1. **`films.xlsx` 永远不要提交**（里面是你的私人数据）。`.gitignore` 已经排除了，但你别手动 `git add -f`。
2. **每次改完先同步再提交**，否则 GitHub Desktop 看不到改动。
3. **版本号到处保持一致**：`worker.py`、页面角标、`版本更新.txt`、README、zip 文件名、Release tag。
4. 想改仓库可见性：仓库页 → Settings → 拉到底 Danger Zone → Change repository visibility。
