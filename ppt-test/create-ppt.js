const pptxgen = require('pptxgenjs');
const html2pptx = require('C:/Users/{替换为用户名}/.claude/skills/pptx-skills/scripts/html2pptx.js');
const path = require('path');

async function createEVAPresentation() {
    const pptx = new pptxgen();

    // 设置演示文稿属性
    pptx.layout = 'LAYOUT_16x9';
    pptx.title = '新世纪福音战士 - 介绍';
    pptx.author = 'Claude Code';
    pptx.subject = 'EVA 动漫介绍';

    const slidesDir = path.join(__dirname);
    const imagesDir = path.join(slidesDir, 'images');

    // 幻灯片文件列表
    const slides = [
        'slide1-title.html',
        'slide2-story.html',
        'slide3-characters-noimg.html',  // 使用无图片版本
        'slide4-eva-units.html',
        'slide5-themes.html',
        'slide6-timeline.html',
        'slide7-impact.html',
        'slide8-thanks.html'
    ];

    // 角色图片配置 (位置单位: 英寸)
    // v3布局: left=50pt, top=85pt, card_width=143pt, gap=15pt, padding=10pt, img_area=123x82pt
    // 转换: 1pt = 1/72 inch
    // 图片x = (50 + 10 + n*(143+15)) / 72, 图片y = (85 + 10) / 72 = 1.32
    // 图片w = 123/72 = 1.71, 图片h = 82/72 = 1.14
    const characterImages = [
        { file: 'shinji.png', x: 0.83, y: 1.32, w: 1.71, h: 1.14 },
        { file: 'rei.png', x: 3.03, y: 1.32, w: 1.71, h: 1.14 },
        { file: 'asuka.png', x: 5.22, y: 1.32, w: 1.71, h: 1.14 },
        { file: 'kaworu.png', x: 7.42, y: 1.32, w: 1.71, h: 1.14 }
    ];

    console.log('Creating EVA Presentation...\n');

    let characterSlide = null;

    // 逐个处理幻灯片
    for (let i = 0; i < slides.length; i++) {
        const slideFile = slides[i];
        const slidePath = path.join(slidesDir, slideFile);

        console.log(`Processing slide ${i + 1}/${slides.length}: ${slideFile}`);

        try {
            const result = await html2pptx(slidePath, pptx);

            // 保存角色页面的 slide 引用
            if (slideFile === 'slide3-characters-noimg.html') {
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
            console.log(`  ✓ Added ${img.file}`);
        }
    }

    // 保存演示文稿
    const outputPath = path.join(slidesDir, 'EVA-Introduction-CN-v3.pptx');
    await pptx.writeFile({ fileName: outputPath });

    console.log(`\n✓ Presentation saved to: ${outputPath}`);
    console.log('\nPresentation Details:');
    console.log('  - Title: 新世纪福音战士 介绍');
    console.log('  - Slides: 8');
    console.log('  - Style: 二次元/赛博朋克风格 NERV配色 (v3优化版)');
    console.log('  - Character images: 4');
}

createEVAPresentation().catch(error => {
    console.error('Failed to create presentation:', error);
    process.exit(1);
});
