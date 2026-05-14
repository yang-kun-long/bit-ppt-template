#!/usr/bin/env node
import pptxgen from 'pptxgenjs';

const pres = new pptxgen();
const slide = pres.addSlide();

// 添加两个矩形作为参考
slide.addShape('rect', {
  x: 1, y: 1, w: 2, h: 1,
  fill: { color: 'E8F4F8' },
  line: { color: '0288D1', width: 2 }
});

slide.addShape('rect', {
  x: 5, y: 3, w: 2, h: 1,
  fill: { color: 'FFF3E0' },
  line: { color: 'F57C00', width: 2 }
});

// 尝试用 line 方法画一条带箭头的线
// 从第一个矩形右边到第二个矩形左边
slide.addShape('line', {
  x: 3,      // 起点 x
  y: 1.5,    // 起点 y
  w: 2,      // 线的宽度（水平长度）
  h: 1.5,    // 线的高度（垂直长度）
  line: {
    color: 'FF0000',
    width: 3,
    endArrowType: 'triangle'  // 尾部箭头
  }
});

// 再画一条弯曲的线（用 curve）
slide.addShape('curve', {
  x: 3,
  y: 2,
  w: 2,
  h: 2,
  line: {
    color: '0000FF',
    width: 3,
    endArrowType: 'triangle'
  }
});

await pres.writeFile({ fileName: 'output/line-arrow-demo.pptx' });
console.log('✅ 生成了 output/line-arrow-demo.pptx');
console.log('请在 PowerPoint/WPS 中打开查看！');
