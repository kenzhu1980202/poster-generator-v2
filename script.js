// 等待DOM加载完成后执行
window.onload = () => {
  // 获取页面上的元素
  const excelFileInput = document.getElementById('excelFile');
  const offerImagesInput = document.getElementById('offerImages');
  const batchGenerateBtn = document.getElementById('batchGenerateBtn');
  const statusDiv = document.getElementById('status');
  const canvas = document.getElementById('posterCanvas');
  const ctx = canvas.getContext('2d');

  // 设置canvas的尺寸
  canvas.width = 750;
  canvas.height = 1334;

  // 点击“批量生成”按钮的事件
  batchGenerateBtn.addEventListener('click', async () => {
    const excelFile = excelFileInput.files[0];
    const imageFiles = offerImagesInput.files;

    if (!excelFile || imageFiles.length === 0) {
      alert('请确保已选择Excel文件和至少一张Offer图片！');
      return;
    }

    // 更新状态并禁用按钮
    statusDiv.innerHTML = '正在初始化...';
    batchGenerateBtn.disabled = true;
    batchGenerateBtn.innerText = '生成中...';

    try {
      // 1. 解析Excel文件
      statusDiv.innerHTML = '正在读取Excel文件...';
      const data = await readExcel(excelFile);
      
      // 2. 创建图片文件的快速查找表
      const imageMap = new Map();
      for (const file of imageFiles) {
        imageMap.set(file.name, file);
      }

      // 3. 初始化JSZip
      const zip = new JSZip();
      const bgImage = await loadImage('background.png'); // 预加载背景图

      // 4. 遍历Excel数据，生成海报
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const school = row.school;
        const major = row.major;
        const imageName = row.image;

        // 更新状态
        statusDiv.innerHTML = `正在处理第 ${i + 1} / ${data.length} 条数据: ${school}`;
        
        // 检查数据完整性
        if (!school || !major || !imageName) {
            console.warn(`第 ${i + 1} 行数据不完整，已跳过。`);
            continue;
        }

        // 查找对应的图片
        const offerImageFile = imageMap.get(imageName);
        if (!offerImageFile) {
            console.warn(`未在您选择的图片中找到名为 "${imageName}" 的文件，已跳过 "${school}"。`);
            continue;
        }
        
        // 生成海报
        const posterDataUrl = await createPoster(bgImage, school, major, offerImageFile);
        
        // 将生成的海报添加到zip中 (需要去掉base64的头部)
        const base64Data = posterDataUrl.split(',')[1];
        zip.file(`喜报-${school}-${major}.png`, base64Data, { base64: true });
      }

      // 5. 生成并下载ZIP文件
      statusDiv.innerHTML = '所有图片处理完毕，正在生成ZIP压缩包...';
      const content = await zip.generateAsync({ type: "blob" });
      
      // 创建下载链接
      const link = document.createElement('a');
      link.href = URL.createObjectURL(content);
      link.download = '批量生成的海报.zip';
      link.click();
      
      statusDiv.innerHTML = '任务完成！ZIP文件已开始下载。';

    } catch (error) {
      console.error('批量生成过程中发生错误:', error);
      statusDiv.innerHTML = `发生错误: ${error.message}`;
      alert('批量生成过程中发生错误，请打开开发者工具(F12)查看Console获取详情。');
    } finally {
      // 恢复按钮状态
      batchGenerateBtn.disabled = false;
      batchGenerateBtn.innerText = '开始批量生成';
    }
  });

  // --- 辅助函数 ---

  // 读取Excel文件并返回JSON数据
  function readExcel(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(worksheet);
          resolve(json);
        } catch (e) {
          reject(e);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  // 核心绘图函数
  async function createPoster(bgImage, school, major, offerImageFile) {
    const offerImage = await loadImage(URL.createObjectURL(offerImageFile));
    
    // 清空并绘制背景
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(bgImage, 0, 0, canvas.width, canvas.height);
    
    // --- 在这里粘贴你之前微调好的绘图代码 ---
    ctx.textAlign = 'left'; 
    ctx.font = 'bold 40px "Microsoft YaHei", sans-serif';
    ctx.fillStyle = '#FFFFFF';
    ctx.fillText(school, 85, 550);
    ctx.font = 'bold 40px "Microsoft YaHei", sans-serif';
    ctx.fillStyle = '#FFFFFF';
    ctx.fillText(major, 85, 670);
    ctx.drawImage(offerImage, 85, 800, 580, 400); 
    // --- 绘图代码结束 ---

    return canvas.toDataURL('image/png');
  }

  // 加载图片的Promise封装
  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'Anonymous';
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = src;
    });
  }
};
