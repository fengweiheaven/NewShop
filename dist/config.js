/* OrderFlow 运行配置。
 * mode: 'local'     —— 单机模式，数据存浏览器 IndexedDB（默认）
 *       'cloudbase' —— 联网模式，数据存腾讯云开发 CloudBase
 * 修改后刷新页面即可生效。云端模式需要 envId（环境 ID）与 accessKey（Publishable Key）。
 */
window.ORDERFLOW_CONFIG = {
  mode: 'local',
  envId: 'orderflow-prod-d2g6wsnki20b1a559',
  region: 'ap-shanghai',
  accessKey: '',
  collection: 'workspaces',
  docId: 'main'
};