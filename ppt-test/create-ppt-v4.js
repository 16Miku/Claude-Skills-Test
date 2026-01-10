const pptxgen = require('pptxgenjs');
const html2pptx = require('C:/Users/{替换为用户名}/.claude/skills/pptx-skills/scripts/html2pptx.js');
const path = require('path');

async function createEVAPresentation() {
    const pptx = new pptxgen();

    // 设置演示文稿属性
    pptx.layout = 'LAYOUT_16x9';
    pptx.title = '新世纪福音战士 - 介绍';
    pptx.author = 'Claude Code';
    pptx.subject = 'EVA 动漫介绍 v4';

    const slidesDir = path.join(__dirname);
    const imagesDir = path.join(slidesDir, 'images');

    // v4 幻灯片文件列表
    const slides = [
        'slide1-title-v4.html',
        'slide2-story-v4.html',
        'slide3-characters-v4.html',  // v4 角色页
        'slide4-eva-units-v4.html',
        'slide5-themes-v4.html',
        'slide6-timeline-v4.html',
        'slide7-impact-v4.html',
        'slide8-thanks-v4.html'
    ];

    // v4 角色图片配置 (位置单位: 英寸)
    // v4布局: left=50pt, top=95pt, card_width=150pt, gap=6pt
    // 图片区域: 宽150pt, 高100pt, 无内边距(图片撑满卡片顶部)
    // 转换: 1pt = 1/72 inch
    // 图片x = (50 + n*(150+6)) / 72
    // 图片y = 95 / 72 = 1.32
    // 图片w = 150/72 = 2.08, 图片h = 100/72 = 1.39
    const characterImages = [
        { file: 'shinji.png', x: 0.69, y: 1.32, w: 2.08, h: 1.39 },  // 50/72
        { file: 'rei.png', x: 2.86, y: 1.32, w: 2.08, h: 1.39 },     // (50+156)/72
        { file: 'asuka.png', x: 5.03, y: 1.32, w: 2.08, h: 1.39 },   // (50+312)/72
        { file: 'kaworu.png', x: 7.19, y: 1.32, w: 2.08, h: 1.39 }   // (50+468)/72
    ];

    console.log('Creating EVA Presentation v4...\n');

    let characterSlide = null;

    // 逐个处理幻灯片
    for (let i = 0; i < slides.length; i++) {
        const slideFile = slides[i];
        const slidePath = path.join(slidesDir, slideFile);

        console.log(`Processing slide ${i + 1}/${slides.length}: ${slideFile}`);

        try {
            const result = await html2pptx(slidePath, pptx);

            // 保存角色页面的 slide 引用
            if (slideFile === 'slide3-characters-v4.html') {
                characterSlide = result.slide;
            }

            console.log(`  ✓ Slide ${i + 1} created successfully`);
        } catch (error) {
            console.error(`  ✗ Error creating slide ${i + 1}:`, error.message);
            throw error;
        }
    }

    // 在角色页面添加图片
    if (characterSlide) {
        console.log('\nAdding character images to slide 3...');
        for (const img of characterImages) {
            const imgPath = path.join(imagesDir, img.file);
            characterSlide.addImage({
                path: imgPath,
                x: img.x,
                y: img.y,
                w: img.w,
                h: img.h
            });
            console.log(`  ✓ Added ${img.file} at (${img.x}, ${img.y})`);
        }
    }

    // 保存演示文稿
    const outputPath = path.join(slidesDir, 'EVA-Introduction-CN-v4.pptx');
    await pptx.writeFile({ fileName: outputPath });

    console.log(`\n✓ Presentation saved to: ${outputPath}`);
    console.log('\nPresentation Details:');
    console.log('  - Title: 新世纪福音战士 介绍');
    console.log('  - Slides: 8');
    console.log('  - Style: 专业渐变设计 + NERV美学 (v4)');
    console.log('  - Character images: 4');
}

createEVAPresentation().catch(error => {
    console.error('Failed to create presentation:', error);
    process.exit(1);
});
