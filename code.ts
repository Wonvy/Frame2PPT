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
}

interface ImageElementData extends ElementData {
  type: 'IMAGE';
  imageHash: string;
  imageBytes: Uint8Array;
  scaleMode?: string;
}

// This shows the HTML page in "ui.html".
figma.showUI(__html__);

// Calls to "parent.postMessage" from within the HTML page will trigger this
// callback. The callback will be passed the "pluginMessage" property of the
// posted message.
figma.ui.onmessage = async (msg: {type: string, count?: number}) => {
  // One way of distinguishing between different types of messages sent from
  // your HTML page is to use an object with a "type" property like this.
  if (msg.type === 'create-shapes' && msg.count !== undefined) {
    // This plugin creates rectangles on the screen.
    const numberOfRectangles = msg.count;

    const nodes: SceneNode[] = [];
    for (let i = 0; i < numberOfRectangles; i++) {
      const rect = figma.createRectangle();
      rect.x = i * 150;
      rect.fills = [{ type: 'SOLID', color: { r: 1, g: 0.5, b: 0 } }];
      figma.currentPage.appendChild(rect);
      nodes.push(rect);
    }
    figma.currentPage.selection = nodes;
    figma.viewport.scrollAndZoomIntoView(nodes);
  }

  if (msg.type === 'export-to-ppt') {
    const selection = figma.currentPage.selection;
    
    if (selection.length === 0) {
      figma.notify('请先选择要导出的Frame');
      return;
    }

    try {
      const frames = selection.filter(node => node.type === 'FRAME');
      
      if (frames.length === 0) {
        figma.notify('请选择至少一个Frame进行导出');
        return;
      }

      const firstFrame = frames[0];
      const aspectRatio = {
        width: firstFrame.width,
        height: firstFrame.height
      };

      const exportPromises = frames.map(async (frame) => {
        const elements: (TextElementData | ShapeElementData | ImageElementData)[] = [];
        
        async function processNode(node: SceneNode, parentX = 0, parentY = 0): Promise<void> {
          // 计算相对于Frame的位置
          const absoluteX = parentX + node.x;
          const absoluteY = parentY + node.y;
          
          // 基本属性
          const baseProps = {
            name: node.name,
            x: absoluteX,
            y: absoluteY,
            width: node.width,
            height: node.height,
            opacity: 'opacity' in node ? (node.opacity as number) : 1,
            rotation: 'rotation' in node ? (node.rotation as number) : 0,
            blendMode: 'blendMode' in node ? node.blendMode : 'NORMAL'
          };

          if (node.type === 'TEXT') {
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
            elements.push(textElement);
          } 
          else if ('fills' in node && 'strokes' in node && 'strokeWeight' in node) {
            // 检查是否包含图片填充
            const nodeWithFills = node as SceneNode & { fills: readonly Paint[] };
            const imageFill = Array.isArray(nodeWithFills.fills) ? 
              nodeWithFills.fills.find(fill => fill.type === 'IMAGE') : 
              null;

            if (imageFill && imageFill.type === 'IMAGE' && imageFill.imageHash) {
              try {
                // 导出图片数据
                const bytes = await node.exportAsync({
                  format: 'PNG',
                  constraint: { type: 'SCALE', value: 2 }
                });

                const imageElement: ImageElementData = {
                  ...baseProps,
                  type: 'IMAGE',
                  imageHash: imageFill.imageHash,
                  imageBytes: bytes,
                  scaleMode: 'FILL'
                };
                elements.push(imageElement);
              } catch (error) {
                console.error('Error exporting image:', error);
                // 如果图片导出失败，作为形状处理
                const geometryNode = node as GeometryMixin;
                const strokeWeight = typeof geometryNode.strokeWeight === 'number' ? geometryNode.strokeWeight : 0;
                
                // 处理圆角
                let cornerRadius: number | undefined = undefined;
                if ('cornerRadius' in node) {
                  const radius = (node as RectangleNode).cornerRadius;
                  if (typeof radius === 'number') {
                    cornerRadius = radius;
                  }
                }

                const shapeElement: ShapeElementData = {
                  ...baseProps,
                  type: 'SHAPE',
                  fills: Array.isArray(nodeWithFills.fills) ? [...nodeWithFills.fills] : [],
                  strokes: Array.isArray(geometryNode.strokes) ? [...geometryNode.strokes] : [],
                  strokeWeight,
                  cornerRadius
                };
                elements.push(shapeElement);
              }
            } else {
              // 不是图片，作为普通形状处理
              const geometryNode = node as GeometryMixin;
              const strokeWeight = typeof geometryNode.strokeWeight === 'number' ? geometryNode.strokeWeight : 0;
              
              // 处理圆角
              let cornerRadius: number | undefined = undefined;
              if ('cornerRadius' in node) {
                const radius = (node as RectangleNode).cornerRadius;
                if (typeof radius === 'number') {
                  cornerRadius = radius;
                }
              }

              const shapeElement: ShapeElementData = {
                ...baseProps,
                type: 'SHAPE',
                fills: Array.isArray(nodeWithFills.fills) ? [...nodeWithFills.fills] : [],
                strokes: Array.isArray(geometryNode.strokes) ? [...geometryNode.strokes] : [],
                strokeWeight,
                cornerRadius
              };
              elements.push(shapeElement);
            }
          }

          // 递归处理子节点
          if ('children' in node) {
            for (const child of node.children) {
              await processNode(child, absoluteX, absoluteY);
            }
          }
        }

        // 处理frame中的所有图层
        await processNode(frame);

        return {
          name: frame.name,
          width: frame.width,
          height: frame.height,
          elements: elements
        };
      });

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

  // Make sure to close the plugin when you're done. Otherwise the plugin will
  // keep running, which shows the cancel button at the bottom of the screen.
  if (msg.type === 'cancel') {
    figma.closePlugin();
  }
};
