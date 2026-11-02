/* OrderFlow 运行配置。
 * mode: 'local'     —— 单机模式，数据存浏览器 IndexedDB（默认，行为与之前一致）
 *       'cloudbase' —— 联网模式，数据存腾讯云开发 CloudBase，需要填 envId
 * 修改后刷新页面即可生效。
 */
window.ORDERFLOW_CONFIG = {
  mode: 'local',
  envId: '',
  collection: 'workspaces',
  docId: 'main'
};
