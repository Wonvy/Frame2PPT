// This plugin will open a window to prompt the user to enter a number, and
// it will then create that many rectangles on the screen.

// This file holds the main code for plugins. Code in this file has access to
// the *figma document* via the figma global object.
// You can access browser APIs in the <script> tag inside "ui.html" which has a
// full browser environment (See https://www.figma.com/plugin-docs/how-plugins-run).

interface ExportData {
  name: string;
  width: number;
  height: number;
  elements: ElementData[];
}

interface ElementData {
  type: string;
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  opacity: number;
  blendMode?: string;
  rotation?: number;
}

interface TextElementData extends ElementData {
  type: 'TEXT';
  characters: string;
  fontSize: number;
  fontName: FontName;
  textAlignHorizontal: string;
  textAlignVertical: string;
  fills: Paint[];
  lineHeight: {
    value: number;
    unit: string;
  } | number;
  letterSpacing: {
    value: number;
    unit: string;
  } | number;
  textCase: string;
  textDecoration: string;
  paragraphIndent: number;
  paragraphSpacing: number;
  textAutoResize: string;
}

interface ShapeElementData extends ElementData {
  type: 'SHAPE';
  fills: Paint[];
  strokes: Paint[];
  strokeWeight: number;
  cornerRadius?: number;
  shapeType?: string;
}

interface ImageElementData extends ElementData {
  type: 'IMAGE';
  imageHash: string;
  imageBytes: Uint8Array;
  scaleMode?: string;
}

interface LineElementData extends ElementData {
  type: 'LINE';
  strokeWeight: number;
  strokes: Paint[];
  strokeCap?: StrokeCap;
  strokeJoin?: StrokeJoin;
  dashPattern?: readonly number[];
  startPoint: { x: number; y: number };
  endPoint: { x: number; y: number };
}

// This shows the HTML page in "ui.html".
figma.showUI(__html__);

// Calls to "parent.postMessage" from within the HTML page will trigger this
// callback. The callback will be passed the "pluginMessage" property of the
// posted message.

// 在文件开头添加一个全局变量来存储当前处理的Frame和元素
let currentFrame: FrameNode;
let currentElements: (TextElementData | ShapeElementData | ImageElementData | LineElementData)[] = [];

async function processFrame(frame: FrameNode): Promise<ExportData> {
  currentFrame = frame;
  currentElements = [];
  
  // 处理Frame中的所有图层
  for (const child of frame.children) {
    await processNode(child);
  }

  return {
    name: frame.name,
    width: frame.width,
    height: frame.height,
    elements: currentElements
  };
}

async function processNode(node: SceneNode): Promise<void> {
  // 计算相对于Frame的位置
  const relativeX = node.absoluteTransform[0][2] - currentFrame.absoluteTransform[0][2];
  const relativeY = node.absoluteTransform[1][2] - currentFrame.absoluteTransform[1][2];
  
  // 基本属性
  const baseProps = {
    name: node.name,
    x: relativeX,
    y: relativeY,
    width: node.width,
    height: node.height,
    opacity: 'opacity' in node ? (node.opacity as number) : 1,
    rotation: 'rotation' in node ? (node.rotation as number) : 0,
    blendMode: 'blendMode' in node ? node.blendMode : 'NORMAL'
  };

  // 处理不同类型的节点
  switch (node.type) {
    case 'INSTANCE':
    case 'COMPONENT': {
      // 检查组件是否包含图片填充
      if ('fills' in node && Array.isArray(node.fills) && node.fills.length > 0) {
        const imageFill = node.fills.find(fill => fill.type === 'IMAGE');
        if (imageFill && imageFill.type === 'IMAGE') {
          try {
            const bytes = await node.exportAsync({
              format: 'PNG',
              constraint: { type: 'SCALE', value: 2 }
            });

            const imageElement: ImageElementData = {
              ...baseProps,
              type: 'IMAGE',
              imageHash: Date.now().toString(),
              imageBytes: bytes,
              scaleMode: 'FILL'
            };
            currentElements.push(imageElement);
            return;
          } catch (error) {
            console.error('Error exporting component with image fill:', error);
          }
        }
      }

      // 处理组件的子元素
      if ('children' in node) {
        for (const child of node.children) {
          await processNode(child);
        }
      }
      return;
    }

    case 'GROUP': {
      if ('children' in node) {
        for (const child of node.children) {
          await processNode(child);
        }
      }
      return;
    }

    case 'FRAME': {
      // 检查Frame是否包含图片填充
      if ('fills' in node && Array.isArray(node.fills) && node.fills.length > 0) {
        const imageFill = node.fills.find(fill => fill.type === 'IMAGE');
        if (imageFill && imageFill.type === 'IMAGE') {
          try {
            const bytes = await node.exportAsync({
              format: 'PNG',
              constraint: { type: 'SCALE', value: 2 }
            });

            const imageElement: ImageElementData = {
              ...baseProps,
              type: 'IMAGE',
              imageHash: Date.now().toString(),
              imageBytes: bytes,
              scaleMode: 'FILL'
            };
            currentElements.push(imageElement);
          } catch (error) {
            console.error('Error exporting frame with image fill:', error);
            await processAsShape(node, baseProps);
          }
        } else {
          // 如果有其他类型的填充，作为形状处理
          await processAsShape(node, baseProps);
        }
      }
      
      // 处理子元素
      if ('children' in node) {
        for (const child of node.children) {
          await processNode(child);
        }
      }
      return;
    }

    case 'TEXT': {
      const textNode = node as TextNode;
      const fontSize = typeof textNode.fontSize === 'number' ? textNode.fontSize : 12;
      
      // 处理行高
      let lineHeight: number | { value: number; unit: string };
      if (typeof textNode.lineHeight === 'number') {
        lineHeight = textNode.lineHeight;
      } else if (textNode.lineHeight && typeof textNode.lineHeight === 'object' && 'value' in (textNode.lineHeight as any)) {
        const lh = textNode.lineHeight as { value: number; unit: string };
        lineHeight = {
          value: lh.value,
          unit: lh.unit
        };
      } else {
        lineHeight = { value: fontSize, unit: 'PIXELS' };
      }

      // 处理字间距
      let letterSpacing: number | { value: number; unit: string };
      if (typeof textNode.letterSpacing === 'number') {
        letterSpacing = textNode.letterSpacing;
      } else if (textNode.letterSpacing && typeof textNode.letterSpacing === 'object' && 'value' in (textNode.letterSpacing as any)) {
        const ls = textNode.letterSpacing as { value: number; unit: string };
        letterSpacing = {
          value: ls.value,
          unit: ls.unit
        };
      } else {
        letterSpacing = { value: 0, unit: 'PIXELS' };
      }

      const textElement: TextElementData = {
        ...baseProps,
        type: 'TEXT',
        characters: textNode.characters,
        fontSize: fontSize,
        fontName: typeof textNode.fontName === 'symbol' ? { family: 'Arial', style: 'Regular' } : textNode.fontName,
        textAlignHorizontal: textNode.textAlignHorizontal,
        textAlignVertical: textNode.textAlignVertical,
        fills: Array.isArray(textNode.fills) ? [...textNode.fills] : [],
        lineHeight,
        letterSpacing,
        textCase: typeof textNode.textCase === 'string' ? textNode.textCase : 'ORIGINAL',
        textDecoration: typeof textNode.textDecoration === 'string' ? textNode.textDecoration : 'NONE',
        paragraphIndent: textNode.paragraphIndent || 0,
        paragraphSpacing: textNode.paragraphSpacing || 0,
        textAutoResize: textNode.textAutoResize
      };
      currentElements.push(textElement);
      return;
    }

    case 'LINE': {
      const lineNode = node as LineNode;
      const strokeWeight = typeof lineNode.strokeWeight === 'number' ? lineNode.strokeWeight : 1;

      // 计算线条的起点和终点（相对于Frame）
      const startX = lineNode.x;
      const startY = lineNode.y;
      const endX = startX + lineNode.width;
      const endY = startY + lineNode.height;

      const lineElement: LineElementData = {
        ...baseProps,
        type: 'LINE',
        strokeWeight: strokeWeight,
        strokes: Array.isArray(lineNode.strokes) ? [...lineNode.strokes] : [],
        strokeCap: typeof lineNode.strokeCap !== 'symbol' ? lineNode.strokeCap : 'NONE',
        strokeJoin: typeof lineNode.strokeJoin !== 'symbol' ? lineNode.strokeJoin : 'MITER',
        dashPattern: lineNode.dashPattern,
        startPoint: {
          x: startX,
          y: startY
        },
        endPoint: {
          x: endX,
          y: endY
        }
      };
      currentElements.push(lineElement);
      return;
    }

    case 'STAR':
    case 'POLYGON': {
      try {
        const bytes = await node.exportAsync({
          format: 'PNG',
          constraint: { type: 'SCALE', value: 2 }
        });

        const imageElement: ImageElementData = {
          ...baseProps,
          type: 'IMAGE',
          imageHash: Date.now().toString(),
          imageBytes: bytes,
          scaleMode: 'FILL'
        };
        currentElements.push(imageElement);
      } catch (error) {
        console.error('Error exporting as image:', error);
        await processAsShape(node, baseProps);
      }
      return;
    }

    case 'RECTANGLE':
    case 'ELLIPSE': {
      // 检查是否包含图片填充
      if ('fills' in node && Array.isArray(node.fills) && node.fills.length > 0) {
        const imageFill = node.fills.find(fill => fill.type === 'IMAGE');
        if (imageFill && imageFill.type === 'IMAGE') {
          try {
            const bytes = await node.exportAsync({
              format: 'PNG',
              constraint: { type: 'SCALE', value: 2 }
            });

            const imageElement: ImageElementData = {
              ...baseProps,
              type: 'IMAGE',
              imageHash: Date.now().toString(),
              imageBytes: bytes,
              scaleMode: 'FILL'
            };
            currentElements.push(imageElement);
            return;
          } catch (error) {
            console.error('Error exporting shape with image fill:', error);
          }
        }
      }
      await processAsShape(node, baseProps);
      return;
    }
  }

  // 处理其他类型的节点
  if ('fills' in node && 'strokes' in node && 'strokeWeight' in node) {
    const nodeWithFills = node as SceneNode & { fills: readonly Paint[] };
    const imageFill = Array.isArray(nodeWithFills.fills) ? 
      nodeWithFills.fills.find(fill => fill.type === 'IMAGE') : 
      null;

    if (imageFill && imageFill.type === 'IMAGE') {
      try {
        const bytes = await node.exportAsync({
          format: 'PNG',
          constraint: { type: 'SCALE', value: 2 }
        });

        const imageElement: ImageElementData = {
          ...baseProps,
          type: 'IMAGE',
          imageHash: Date.now().toString(),
          imageBytes: bytes,
          scaleMode: 'FILL'
        };
        currentElements.push(imageElement);
      } catch (error) {
        console.error('Error exporting image:', error);
        await processAsShape(node, baseProps);
      }
    } else {
      await processAsShape(node, baseProps);
    }
  }
}

async function processAsShape(node: SceneNode, baseProps: any) {
  if (!('fills' in node) || !('strokes' in node) || !('strokeWeight' in node)) {
    return;
  }

  const geometryNode = node as GeometryMixin;
  const strokeWeight = typeof geometryNode.strokeWeight === 'number' ? geometryNode.strokeWeight : 0;
  
  let cornerRadius: number | undefined = undefined;
  if ('cornerRadius' in node) {
    const radius = (node as RectangleNode).cornerRadius;
    if (typeof radius === 'number') {
      cornerRadius = radius;
    }
  }

  // 确定形状类型
  let shapeType = 'RECTANGLE';
  if (node.type === 'ELLIPSE') {
    shapeType = 'ELLIPSE';
  } else if (node.type === 'POLYGON') {
    shapeType = 'POLYGON';
  } else if (node.type === 'STAR') {
    shapeType = 'STAR';
  }

  const shapeElement: ShapeElementData = {
    ...baseProps,
    type: 'SHAPE',
    fills: Array.isArray(node.fills) ? [...node.fills] : [],
    strokes: Array.isArray(geometryNode.strokes) ? [...geometryNode.strokes] : [],
    strokeWeight,
    cornerRadius,
    shapeType
  };
  currentElements.push(shapeElement);
}

// 修改主导出函数
figma.ui.onmessage = async (msg: {type: string, count?: number}) => {
  if (msg.type === 'export-to-ppt') {
    const selection = figma.currentPage.selection;
    
    if (selection.length === 0) {
      figma.notify('请先选择要导出的Frame');
      return;
    }

    try {
      const frames = selection.filter(node => node.type === 'FRAME') as FrameNode[];
      
      if (frames.length === 0) {
        figma.notify('请选择至少一个Frame进行导出');
        return;
      }

      const firstFrame = frames[0];
      const aspectRatio = {
        width: firstFrame.width,
        height: firstFrame.height
      };

      const exportPromises = frames.map(frame => processFrame(frame));
      const exportedFrames = await Promise.all(exportPromises);
      
      figma.ui.postMessage({
        type: 'export-data',
        data: exportedFrames,
        aspectRatio: aspectRatio
      });
      
      figma.notify('导出成功！');
    } catch (error: any) {
      figma.notify('导出失败：' + error.message);
    }
  }

  if (msg.type === 'cancel') {
    figma.closePlugin();
  }
};
