import type { MarkdownImageAdapter } from 'markdown-docx'
import markdownDocx, { MarkdownDocx, Packer } from 'markdown-docx'
import * as echarts from 'echarts'

type EChartImageData = {
  index: number
  imageData: string
  width?: number
  height?: number
}

/**
 * Calculate target dimensions for Word document while maintaining aspect ratio
 * @param width Original image width
 * @param height Original image height
 * @param maxWidth Maximum width (default: 600)
 * @param maxHeight Maximum height (default: 400)
 * @returns Target dimensions { width, height }
 */
function calculateTargetDimensions(
  width: number,
  height: number,
  maxWidth: number = 600,
  maxHeight: number = 400,
): { width: number; height: number } {
  let targetWidth = width
  let targetHeight = height

  // Scale down if exceeds max dimensions, maintaining aspect ratio
  if (targetWidth > maxWidth) {
    const ratio = maxWidth / targetWidth
    targetWidth = maxWidth
    targetHeight = targetHeight * ratio
  }

  if (targetHeight > maxHeight) {
    const ratio = maxHeight / targetHeight
    targetHeight = maxHeight
    targetWidth = targetWidth * ratio
  }

  return { width: Math.round(targetWidth), height: Math.round(targetHeight) }
}

/**
 * Get all ECharts instance images from the specified DOM container
 * @param containerElement The DOM container of the message content
 */
async function getEChartsImagesFromDOM(containerElement: HTMLElement): Promise<EChartImageData[]> {
  const images: EChartImageData[] = []

  // Find all echarts render containers within the container
  // echarts-for-react sets the _echarts_instance_ attribute on the container by default
  const echartsContainers = containerElement.querySelectorAll('[_echarts_instance_]')

  const imagePromises = Array.from(echartsContainers).map(async (container, index) => {
    try {
      const instance = echarts.getInstanceByDom(container as HTMLElement)
      if (instance) {
        const imageData = instance.getDataURL({
          type: 'png',
          pixelRatio: 2,
          backgroundColor: '#fff',
          excludeComponents: ['toolbox', 'dataZoom'],
        })

        // Get image dimensions and calculate target dimensions
        try {
          const img = new Image()
          await new Promise((resolve, reject) => {
            img.onload = resolve
            img.onerror = reject
            img.src = imageData
          })

          const originalWidth = img.naturalWidth
          const originalHeight = img.naturalHeight
          const targetDimensions = calculateTargetDimensions(originalWidth, originalHeight, 600, 400)

          images.push({
            index,
            imageData,
            width: targetDimensions.width,
            height: targetDimensions.height,
          })
        }
        catch (error) {
          // If image loading fails, add image without dimensions
          console.warn(`Failed to get dimensions for ECharts image at index ${index}:`, error)
          images.push({ index, imageData })
        }
      }
    }
    catch (error) {
      console.warn(`Failed to get ECharts image at index ${index}:`, error)
    }
  })

  await Promise.all(imagePromises)

  return images
}

/**
 * Parse ECharts code block positions in Markdown
 */
function parseEChartsBlockPositions(content: string): Array<{ startIndex: number; endIndex: number }> {
  const positions: Array<{ startIndex: number; endIndex: number }> = []
  const regex = /```echarts\n[\s\S]*?```/g
  let match

  while ((match = regex.exec(content)) !== null) {
    positions.push({
      startIndex: match.index,
      endIndex: match.index + match[0].length,
    })
  }

  return positions
}

/**
 * Replace ECharts code blocks with images
 */
async function replaceEChartsBlocksWithImages(
  content: string,
  blockPositions: Array<{ startIndex: number; endIndex: number }>,
  images: EChartImageData[],
): Promise<string> {
  let result = content

  // Replace from back to front to avoid index offset
  const sortedPositions = [...blockPositions].sort((a, b) => b.startIndex - a.startIndex)

  sortedPositions.forEach((pos, reverseIndex) => {
    // Find the corresponding image (original index)
    const originalIndex = blockPositions.length - 1 - reverseIndex
    const imageData = images.find(img => img.index === originalIndex)

    if (imageData) {
      // If dimensions are available, include them in the markdown
      const imageMarkdown = imageData.width && imageData.height
        ? `![ECharts Chart](${imageData.imageData} "${imageData.width}x${imageData.height}")`
        : `![ECharts Chart](${imageData.imageData})`
      result = result.slice(0, pos.startIndex) + imageMarkdown + result.slice(pos.endIndex)
    }
  })

  return result
}

/**
 * Remove <think>...</think> tags and their content from markdown
 */
function removeThinkTags(content: string): string {
  // Remove <think>...</think> tags (including multiline content)
  return content.replace(/<think>[\s\S]*?<\/think>/gi, '')
}

/**
 * Generate filename
 */
function generateFilename(content: string): string {
  const now = new Date()
  const timestamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`

  // Try to extract title from content
  const titleMatch = content.match(/^#\s+(.+)$/m)
  if (titleMatch) {
    const title = titleMatch[1].slice(0, 30).replace(/[\\/:*?"<>|]/g, '')
    return `${title}_${timestamp}.docx`
  }

  return `chat_export_${timestamp}.docx`
}

/**
 * Convert markdown to docx blob
 * @param markdownText Markdown content
 * @returns Blob of the docx file
 */
async function convertMarkdownToDocx(markdownText: string): Promise<Blob> {
  // Convert to docx
  const imageAdapter: MarkdownImageAdapter = async (token) => {
    const defaultAdapter = MarkdownDocx.defaultOptions.imageAdapter!
    // try to parse width and height from token.title, like ![](image.png "600x400")
    const [width, height] = token?.title?.match(/(\d+)x(\d+)/)?.slice(1) ?? []
    const result = await defaultAdapter(token)
    if (result?.width && result?.height && width && height) {
      return {
        ...result,
        width: Number.parseInt(width, 10),
        height: Number.parseInt(height, 10),
      }
    }
    return result
  }
  const doc = await markdownDocx(markdownText, {
    imageAdapter,
  })
  // Generate blob for download
  const blob = await Packer.toBlob(doc)
  return blob
}

/**
 * Download docx file
 * @param blob Blob of the docx file
 * @param filename Filename for download
 */
function downloadDocx(blob: Blob, filename: string): void {
  // Create download link
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.style.display = 'none'
  document.body.appendChild(a)
  a.click()

  // Clean up immediately - browser saves URL reference on click
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

/**
 * Download Word document
 * @param content Markdown content
 * @param containerElement The DOM container of the message content (used to get ECharts images)
 * @param filename Optional filename
 */
export async function downloadAsDocx(
  content: string,
  containerElement?: HTMLElement | null,
  filename?: string,
): Promise<void> {
  try {
    // Remove <think>...</think> tags first
    let processedContent = removeThinkTags(content)

    // If container element is provided, try to get ECharts images from DOM
    if (containerElement) {
      // Parse block positions from processedContent to ensure position accuracy
      const blockPositions = parseEChartsBlockPositions(processedContent)

      if (blockPositions.length > 0) {
        // Get rendered ECharts images from DOM
        const images = await getEChartsImagesFromDOM(containerElement)

        if (images.length > 0)
          processedContent = await replaceEChartsBlocksWithImages(content, blockPositions, images)
      }
    }

    // Convert to Word document
    const docxBlob = await convertMarkdownToDocx(processedContent)

    // Trigger download
    const finalFilename = filename || generateFilename(content)
    downloadDocx(docxBlob, finalFilename)
  }
  catch (error) {
    console.error('Failed to download as docx:', error)
    throw error
  }
}

export default {
  downloadAsDocx,
}
