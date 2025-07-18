// script.js (最终修复版 V3 - 恢复文章大标题)

window.onload = () => {
    // --- DOM 元素获取 (无变化) ---
    const uploadView = document.getElementById('upload-view');
    const configView = document.getElementById('config-view');
    const statusView = document.getElementById('status-view');
    const fileInput = document.getElementById('wordFiles');
    const fileInputWrapper = document.querySelector('.file-input-wrapper');
    const fileLabelText = document.getElementById('file-label-text');
    const configFormsContainer = document.getElementById('config-forms-container');
    const processAllBtn = document.getElementById('processAllBtn');
    const statusLog = document.getElementById('status-log');

    // --- 全局状态变量 (无变化) ---
    let uploadedFilesData = [];

    // --- 事件监听 (无变化) ---
    fileInput.addEventListener('change', handleFileSelect);
    processAllBtn.addEventListener('click', handleProcessAll);
    fileInputWrapper.addEventListener('dragover', (e) => { e.preventDefault(); fileInputWrapper.style.backgroundColor = '#eef4ff'; fileInputWrapper.style.borderColor = '#1a73e8'; });
    fileInputWrapper.addEventListener('dragleave', (e) => { e.preventDefault(); fileInputWrapper.style.backgroundColor = ''; fileInputWrapper.style.borderColor = ''; });
    fileInputWrapper.addEventListener('drop', (e) => { e.preventDefault(); fileInputWrapper.style.backgroundColor = ''; fileInputWrapper.style.borderColor = ''; fileInput.files = e.dataTransfer.files; handleFileSelect({ target: fileInput }); });

    // --- 函数定义 ---

    async function handleFileSelect(event) {
        const files = event.target.files;
        if (files.length === 0) return;
        fileLabelText.textContent = `已选择 ${files.length} 个文件`;
        processAllBtn.disabled = true;
        uploadedFilesData = [];
        configFormsContainer.innerHTML = '';
        uploadView.style.display = 'none';
        configView.style.display = 'block';
        statusView.style.display = 'none';
        const cardPromises = [];
        for (let i = 0; i < files.length; i++) {
            const file = files.item(i);
            const fileData = { file, originalName: file.name, images: [] };
            uploadedFilesData.push(fileData);
            cardPromises.push(generateConfigCard(fileData, i));
        }
        await Promise.all(cardPromises);
        processAllBtn.disabled = false;
    }

    async function generateConfigCard(fileData, fileIndex) {
        const card = document.createElement('div');
        card.className = 'config-card';
        card.innerHTML = `<h3>${fileData.originalName}</h3>`;
        configFormsContainer.appendChild(card);
        try {
            const zip = await JSZip.loadAsync(fileData.file);
            const imagePromises = [];
            zip.folder("word/media").forEach((relativePath, imageFile) => {
                if (!imageFile.dir) {
                    const promise = imageFile.async('blob').then(blob => ({ blob, name: imageFile.name, extension: relativePath.split('.').pop() }));
                    imagePromises.push(promise);
                }
            });
            const images = await Promise.all(imagePromises);
            if (images.length === 0) {
                card.innerHTML += `<p>文件中未检测到图片。</p>`;
            } else {
                images.forEach((img, imageIndex) => {
                    const objectURL = URL.createObjectURL(img.blob);
                    fileData.images.push({ blob: img.blob, originalName: img.name, extension: img.extension, altText: '', dimensions: { width: 0, height: 0 } });
                    const imageConfigDiv = document.createElement('div');
                    imageConfigDiv.className = 'image-config-item';
                    imageConfigDiv.innerHTML = `<img src="${objectURL}" alt="${img.name}" class="thumbnail"><input type="text" placeholder="填写ALT文本，例如: a-black-dog" class="alt-input" data-file-index="${fileIndex}" data-image-index="${imageIndex}">`;
                    card.appendChild(imageConfigDiv);
                    const thumbnailImg = imageConfigDiv.querySelector('.thumbnail');
                    thumbnailImg.onload = () => URL.revokeObjectURL(thumbnailImg.src);
                });
            }
        } catch (error) {
            card.innerHTML += `<p style="color: red;">无法读取此文件...</p>`;
            console.error(`Error processing ${fileData.originalName}:`, error);
        }
    }

    async function handleProcessAll() {
        configView.style.display = 'none';
        statusView.style.display = 'block';
        clearStatus();
        processAllBtn.disabled = true;

        logStatus("步骤 1/4: 读取用户输入的ALT文本...");
        document.querySelectorAll('.alt-input').forEach(input => {
            const fileIndex = parseInt(input.dataset.fileIndex);
            const imageIndex = parseInt(input.dataset.imageIndex);
            const altValue = input.value.trim() || uploadedFilesData[fileIndex].images[imageIndex].originalName.split('/').pop().split('.')[0];
            uploadedFilesData[fileIndex].images[imageIndex].altText = altValue;
        });
        logStatus("-> ALT文本读取完毕。");

        logStatus("\n步骤 2/4: 获取所有图片原始尺寸...");
        const dimensionPromises = [];
        uploadedFilesData.forEach(fileData => fileData.images.forEach(img => {
            dimensionPromises.push(getImageDimensions(img.blob).then(dims => { img.dimensions = dims; }));
        }));
        await Promise.all(dimensionPromises);
        logStatus("-> 图片尺寸获取完毕。");

        logStatus("\n步骤 3/4: 开始转换文件并打包...");
        for (const fileData of uploadedFilesData) {
            try {
                logStatus(`\n--- 正在处理: ${fileData.originalName} ---`);
                const zipBlob = await processSingleFile(fileData);
                const zipName = `${fileData.originalName.replace(/\.docx$/, '')}.zip`;
                saveAs(zipBlob, zipName);
                logStatus(`-> 成功创建并下载: ${zipName}`);
            } catch (error) {
                logStatus(`-> !!! 处理失败: ${fileData.originalName} - ${error.message}`);
                console.error(`Failed to process ${fileData.originalName}:`, error);
            }
        }
        logStatus("\n步骤 4/4: 所有任务完成！");

        setTimeout(() => {
            uploadView.style.display = 'block';
            statusView.style.display = 'none';
            fileInput.value = '';
            fileLabelText.textContent = '点击此处选择或拖拽 .docx 文件';
        }, 3000);
    }
    
    async function processSingleFile(fileData) {
        logStatus("  - 正在读取文件内容...");
        const arrayBuffer = await fileData.file.arrayBuffer();
        
        let imageCounter = 0;
        
        logStatus("  - [流程] 步骤1: 按顺序生成带标记的初步HTML...");
        const options = {
            convertImage: mammoth.images.imgElement(image => {
                const imgData = fileData.images[imageCounter];
                if (imgData) {
                    const marker = `---IMG-GUID-${imageCounter}---`;
                    imageCounter++;
                    return { alt: marker, src: '' };
                }
                return {};
            })
        };
        let { value: preliminaryHtml } = await mammoth.convertToHtml({ arrayBuffer }, options);
        
        // ===================== H1标题处理：步骤1 开始 =====================
        // 优先提取H1纯文本，并从主HTML中移除H1标签，防止后续处理
        const { value: rawText } = await mammoth.extractRawText({ arrayBuffer });
        const h1Text = (rawText.split('\n')[0] || fileData.originalName.replace(/\.docx$/, '')).trim();
        preliminaryHtml = preliminaryHtml.replace(/<h1>[\s\S]*?<\/h1>/, '');
        logStatus(`  - 已提取文章标题: ${h1Text}`);
        // ===================== H1标题处理：步骤1 结束 =====================


        logStatus("  - [流程] 步骤2: 逐段处理并重建HTML...");
        const blockRegex = /(<(p|h2|h3|ul|ol|table)[\s\S]*?<\/\2>)/g;
        const htmlBlocks = preliminaryHtml.match(blockRegex) || [];
        
        const finalHtmlParts = [];

        for (const block of htmlBlocks) {
            let processedBlock = block;
            const imageMatch = processedBlock.match(/---IMG-GUID-(\d+)---/);
            
            if (imageMatch) {
                logStatus("    - 发现图片段落，正在替换...");
                const index = parseInt(imageMatch[1]);
                const imgData = fileData.images[index];

                if (imgData) {
                    const newFileName = `${createSafeFilename(imgData.altText)}.${imgData.extension}`;
                    const imageHtmlBlock = `<p style="text-align:center"><img alt="${imgData.altText}" src="${newFileName}" style="height:${imgData.dimensions.height}px; width:${imgData.dimensions.width}px" /></p>\n<p style="text-align:center"><em>${imgData.altText}</em></p>`;
                    finalHtmlParts.push(imageHtmlBlock);
                    logStatus(`    - 已插入图片: ${newFileName}`);
                }
            } else {
                if (processedBlock.startsWith('<p>')) {
                    processedBlock = processedBlock.replace(/<\/p>/g, '</p>\n<p>&nbsp;</p>');
                }
                if (processedBlock.startsWith('<h2>')) {
                    processedBlock = processedBlock.replace('<h2>', '<h2 style="margin-top: 24px">');
                }
                if (processedBlock.startsWith('<h3>')) {
                    processedBlock = processedBlock.replace('<h3>', '<h3 style="margin-top: 24px">');
                }
                finalHtmlParts.push(processedBlock);
            }
        }

        let finalHtmlBody = finalHtmlParts.join('\n\n');
        
        logStatus("  - 格式化HTML代码使其美观...");
        const formattedBody = html_beautify(finalHtmlBody, {
            indent_size: 2,
            space_in_empty_paren: true,
            wrap_line_length: 120
        });
        
        // ===================== H1标题处理：步骤2 开始 =====================
        // 将纯文本标题加在格式化好的正文前面
        const finalFileContent = h1Text + '\n\n' + formattedBody;
        // ===================== H1标题处理：步骤2 结束 =====================

        logStatus("  - 正在创建ZIP压缩包...");
        const outputZip = new JSZip();
        const baseFilename = createSafeFilename(h1Text); 

        // 同时创建 .html 和 .txt 文件，它们的内容都是 finalFileContent
        outputZip.file(`${baseFilename}.html`, finalFileContent);
        outputZip.file(`${baseFilename}.txt`, finalFileContent);

        for (const img of fileData.images) {
            const newFileName = `${createSafeFilename(img.altText)}.${img.extension}`;
            outputZip.file(newFileName, img.blob);
        }

        logStatus("  - 生成ZIP文件中...");
        return outputZip.generateAsync({ type: "blob" });
    }

    // --- 帮助函数 (无变化) ---
    function logStatus(message) { statusLog.textContent += message + '\n'; statusLog.scrollTop = statusLog.scrollHeight; }
    function clearStatus() { statusLog.textContent = ''; }
    function getImageDimensions(blob) { return new Promise((resolve, reject) => { const img = new Image(); const objectURL = URL.createObjectURL(blob); img.onload = () => { resolve({ width: img.width, height: img.height }); URL.revokeObjectURL(objectURL); }; img.onerror = (err) => { reject(err); URL.revokeObjectURL(objectURL); }; img.src = objectURL; }); }
    function createSafeFilename(text) { if (!text) return 'untitled'; return text.toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5\s-]/g, '').trim().replace(/\s+/g, '-'); }
};
