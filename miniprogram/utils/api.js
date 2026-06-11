// API 请求封装
const app = getApp();

function request(path, options = {}) {
  const { method = 'GET', data, filePath } = options;
  const base = app.globalData.apiBase;

  return new Promise((resolve, reject) => {
    const params = {
      url: base + path,
      method,
      header: {},
      success: (res) => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(res.data);
        } else {
          reject(new Error(res.data?.error || `请求失败 (${res.statusCode})`));
        }
      },
      fail: (err) => reject(new Error(err.errMsg || '网络错误'))
    };

    if (filePath) {
      params.name = 'file';
      params.filePath = filePath;
    } else if (data && method === 'POST') {
      params.header['Content-Type'] = 'application/json';
      params.data = data;
    }

    if (filePath) {
      wx.uploadFile(params);
    } else {
      wx.request(params);
    }
  });
}

// 上传文件（带真实进度回调，返回 promise + task 引用）
function uploadFile(filePath, onProgress) {
  let task;
  const promise = new Promise((resolve, reject) => {
    task = wx.uploadFile({
      url: app.globalData.apiBase + '/analyze',
      filePath,
      name: 'file',
      success: (res) => {
        try {
          const data = JSON.parse(res.data);
          if (data.success) resolve(data);
          else reject(new Error(data.error));
        } catch (e) {
          reject(new Error('解析响应失败'));
        }
      },
      fail: (err) => reject(new Error(err.errMsg || '上传失败'))
    });

    // 真实上传进度
    if (onProgress) {
      task.onProgressUpdate((res) => {
        onProgress(res.progress); // 0-100
      });
    }
  });
  promise.task = task;
  return promise;
}

// 下载文件（带真实进度回调，返回 promise + task 引用）
function downloadFileWithProgress(url, onProgress) {
  let task;
  const promise = new Promise((resolve, reject) => {
    task = wx.downloadFile({
      url,
      success: (res) => {
        if (res.statusCode === 200) {
          resolve(res);
        } else {
          reject(new Error('下载失败'));
        }
      },
      fail: reject
    });

    // 真实下载进度
    if (onProgress) {
      task.onProgressUpdate((res) => {
        onProgress(res.progress); // 0-100
      });
    }
  });
  promise.task = task;
  return promise;
}

// 下载文件并打开
function downloadFile(url, filename) {
  return new Promise((resolve, reject) => {
    wx.downloadFile({
      url: app.globalData.apiBase + url,
      success: (res) => {
        if (res.statusCode === 200) {
          wx.openDocument({
            filePath: res.tempFilePath,
            showMenu: true,
            success: resolve,
            fail: reject
          });
        } else {
          reject(new Error('下载失败'));
        }
      },
      fail: reject
    });
  });
}

module.exports = { request, uploadFile, downloadFile, downloadFileWithProgress };
