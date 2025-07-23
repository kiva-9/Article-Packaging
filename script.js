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

// [最终修复版 V2] 请用此函数完整替换旧的 generateConfigCard 函数
    async function generateConfigCard(fileData, fileIndex) {
        const card = document.createElement('div');
        card.className = 'config-card';
        card.innerHTML = `<h3>${fileData.originalName}</h3>`;
        configFormsContainer.appendChild(card);

        try {
            const arrayBuffer = await fileData.file.arrayBuffer();
            
            const imagePromises = []; 

            const discoveryOptions = {
                convertImage: mammoth.images.imgElement((image) => {
                    // [FIX] 调用 image.read()时不带参数来获取ArrayBuffer
                    const promise = image.read()
                        .then(imageArrayBuffer => {
                            // [FIX] 手动从ArrayBuffer和contentType创建Blob
                            const blob = new Blob([imageArrayBuffer], { type: image.contentType });
                            const extension = image.contentType.split('/')[1] || 'png';
                            return { blob, extension };
                        });
                    
                    imagePromises.push(promise);
                    
                    return { src: "" };
                })
            };

            await mammoth.convertToHtml({ arrayBuffer }, discoveryOptions);
            
            const discoveredImages = await Promise.all(imagePromises);
            
            if (discoveredImages.length === 0) {
                card.innerHTML += `<p>文件中未检测到图片。</p>`;
            } else {
                discoveredImages.forEach((img, imageIndex) => {
                    const objectURL = URL.createObjectURL(img.blob);
                    const placeholderName = `document-image-${imageIndex + 1}`;
                    
                    fileData.images.push({ 
                        blob: img.blob, 
                        originalName: `${placeholderName}.${img.extension}`, 
                        extension: img.extension, 
                        altText: '', 
                        dimensions: { width: 0, height: 0 } 
                      });
                    
                    const imageConfigDiv = document.createElement('div');
                    imageConfigDiv.className = 'image-config-item';
                    imageConfigDiv.innerHTML = `<img src="${objectURL}" alt="${placeholderName}" class="thumbnail"><input type="text" placeholder="填写ALT文本，例如: a-black-dog" class="alt-input" data-file-index="${fileIndex}" data-image-index="${imageIndex}">`;
                    card.appendChild(imageConfigDiv);
                    
                    const thumbnailImg = imageConfigDiv.querySelector('.thumbnail');
                    thumbnailImg.onload = () => URL.revokeObjectURL(thumbnailImg.src); 
                });
            }
        } catch (error) {
            card.innerHTML += `<p style="color: red;">无法读取此文件... ${error.message}</p>`;
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
    
// [最终修复版 V5 - “魔法词”居中方案] 请用此函数完整替换旧的 processSingleFile 函数
    async function processSingleFile(fileData) {
        logStatus("  - 正在读取文件内容...");
        const arrayBuffer = await fileData.file.arrayBuffer();

        let imageCounter = 0;

        logStatus("  - [流程] 步骤1: 按顺序生成带标记的初步HTML...");

        // [重大调整] 移除所有 styleMap 相关的代码，因为对您的文档无效。
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

        const { value: rawText } = await mammoth.extractRawText({ arrayBuffer });
        const h1Text = (rawText.split('\n')[0] || fileData.originalName.replace(/\.docx$/, '')).trim();
        preliminaryHtml = preliminaryHtml.replace(/<h1>[\s\S]*?<\/h1>/, '');
        logStatus(`  - 已提取文章标题: ${h1Text}`);
        
        logStatus("  - [流程] 步骤2: 逐段处理并重建HTML...");
        const blockRegex = /(<(p|h2|h3|h4|ul|ol|table)[\s\S]*?<\/\2>)/g;
        const htmlBlocks = preliminaryHtml.match(blockRegex) || [];

        const finalHtmlParts = [];

        for (let i = 0; i < htmlBlocks.length; i++) {
            let block = htmlBlocks[i];
            let processedBlock = block;
            const imageMatch = processedBlock.match(/---IMG-GUID-(\d+)---/);

            if (imageMatch) {
                logStatus("  - 发现图片段落，正在应用特殊规则...");
                const index = parseInt(imageMatch[1]);
                const imgData = fileData.images[index];

                if (imgData) {
                    const placeholder = `{{${createSafeFilename(imgData.altText)}}}`;
                    const imageHtmlBlock = `<p>&nbsp;</p>\n<p>&nbsp;</p>\n<p style="text-align:center"><img alt="${imgData.altText}" src="${placeholder}" style="height:${imgData.dimensions.height}px; width:${imgData.dimensions.width}px" /></p>\n<p>&nbsp;</p>\n<p>&nbsp;</p>`;
                    finalHtmlParts.push(imageHtmlBlock);
                    logStatus(`    - 已插入图片占位符: ${placeholder}`);
                }
            } else {
                // [新增] "魔法词"居中方案
                const centerMarker = '[center]';
                if (processedBlock.includes(centerMarker)) {
                    logStatus(`  - 检测到居中标记 "[center]"，正在应用居中样式...`);
                    // 为该块添加居中样式。这是一个简单但有效的实现。
                    processedBlock = processedBlock.replace('>', ' style="text-align: center;">');
                    // 移除魔法词本身，避免它显示在最终内容中
                    processedBlock = processedBlock.replace(centerMarker, '');
                }

                if (processedBlock.startsWith('<p>')) {
                    const nextBlock = htmlBlocks[i + 1];
                    if (nextBlock && nextBlock.startsWith('<p>')) {
                        processedBlock = processedBlock + '\n<p>&nbsp;</p>';
                    }
                }
                
                // [修改] 调整标题处理逻辑，使其能与上面的“魔法词”方案正确配合
                const headingMatch = processedBlock.match(/^<(h[234])/);
                if (headingMatch) {
                    if (processedBlock.includes('style="')) {
                        // 如果已经有style（说明它被上面的魔法词居中了），我们就在里面追加样式
                        processedBlock = processedBlock.replace('style="', `style="margin-top: 24rem; `);
                    } else {
                        // 如果没有style，就创建一个新的
                        const headingTag = headingMatch[1]; 
                        processedBlock = processedBlock.replace(`<${headingTag}>`, `<${headingTag} style="margin-top: 24rem">`);
                    }
                }
                finalHtmlParts.push(processedBlock);
            }
        }

        let finalHtmlBody = finalHtmlParts.join('\n\n');
        
        const threeOrMoreEmptyPs = /(<p>&nbsp;<\/p>\s*){3,}/g;
        const twoEmptyPs = '<p>&nbsp;</p>\n<p>&nbsp;</p>';
        finalHtmlBody = finalHtmlBody.replace(threeOrMoreEmptyPs, twoEmptyPs);

        const formattedBody = html_beautify(finalHtmlBody, {
            indent_size: 2,
            space_in_empty_paren: true,
            wrap_line_length: 120
        });

        const finalFileContent = h1Text + '\n\n' + formattedBody;

        logStatus("  - 正在创建ZIP压缩包...");
        const outputZip = new JSZip();
        const baseFilename = createSafeFilename(h1Text);

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
