// This plugin will open a window to prompt the user to enter a number, and
// it will then create that many rectangles on the screen.

// This file holds the main code for plugins. Code in this file has access to
// the *figma document* via the figma global object.
// You can access browser APIs in the <script> tag inside "ui.html" which has a
// full browser environment (See https://www.figma.com/plugin-docs/how-plugins-run).

interface ExportData {
  name: string;
  bytes: Uint8Array;
  width: number;
  height: number;
  textNodes: TextNodeData[];
}

interface TextNodeData {
  x: number;
  y: number;
  width: number;
  height: number;
  characters: string;
  fontSize: number;
  fontName: FontName;
  textAlignHorizontal: string;
  textAlignVertical: string;
  fills: Paint[];
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
      // 获取所有选中的Frame
      const frames = selection.filter(node => node.type === 'FRAME');
      
      if (frames.length === 0) {
        figma.notify('请选择至少一个Frame进行导出');
        return;
      }

      // 获取第一个Frame的尺寸比例
      const firstFrame = frames[0];
      const aspectRatio = {
        width: firstFrame.width,
        height: firstFrame.height
      };

      // 导出每个Frame的信息
      const exportPromises = frames.map(async (frame) => {
        // 获取Frame中的所有文本节点
        const textNodes: TextNodeData[] = [];
        
        function traverseNode(node: SceneNode) {
          if (node.type === 'TEXT') {
            // 计算相对于Frame的位置
            const relativeX = node.x - frame.x;
            const relativeY = node.y - frame.y;
            
            textNodes.push({
              x: relativeX,
              y: relativeY,
              width: node.width,
              height: node.height,
              characters: node.characters,
              fontSize: node.fontSize,
              fontName: node.fontName,
              textAlignHorizontal: node.textAlignHorizontal,
              textAlignVertical: node.textAlignVertical,
              fills: node.fills as Paint[]
            });
          }
          
          if ('children' in node) {
            (node.children as SceneNode[]).forEach(child => traverseNode(child));
          }
        }
        
        traverseNode(frame);

        // 导出Frame为PNG
        const bytes = await frame.exportAsync({
          format: 'PNG',
          constraint: { type: 'SCALE', value: 2 }
        });
        
        return {
          name: frame.name,
          bytes: bytes,
          width: frame.width,
          height: frame.height,
          textNodes: textNodes
        } as ExportData;
      });

      const exportedImages = await Promise.all(exportPromises);
      
      // 发送导出数据到UI
      figma.ui.postMessage({
        type: 'export-data',
        data: exportedImages,
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
