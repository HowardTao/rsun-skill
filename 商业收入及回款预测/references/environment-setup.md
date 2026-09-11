# guancli 安装与认证（环境搭建，低频参考）

> 本文件为**低频参考**：仅当 `guancli` 未安装、认证失效或权限不足时需要阅读；日常执行流程见 `SKILL.md` 第 0 步。

## 安装

- 前置：Node.js 20+（推荐 22+）。
- 安装命令（更新方式：重新执行同样命令即可）：

```bash
npm install -g --foreground-scripts @guandata/guanskill
guanskill install-skill
```

- 官方文档：https://www.guandata.com/guandata-cli-installation-guide.md（CLI 集合包安装、认证登录、各组件版本检查）。

## 认证登录

```bash
guancli auth login    # 配置观远访问凭证（PAT）
guancli auth status   # 确认 PAT 是否有效
```

- 部分步骤需用户在浏览器中配合完成授权。
- 用户名密码登录时 Domain 默认 `guanbi`，也可留空自动探测。
- 安装/登录命令涉及系统级操作，需用户明确确认后再执行。

## BI 环境与数据集权限

- 目标 BI 环境：https://rsunbi.rsun.com:9521（guancli default profile）。
- 需向观远管理员申请以下两个数据集的访问权限：
  - `ADS-租费分析-租赁合同台账明细`（dsId `k3a8a2772ee4143d694a2ecd`）
  - `ADS-租费分析-合同全周期应收明细`（dsId `ud1f309f6e04d4285b7364c4`）
