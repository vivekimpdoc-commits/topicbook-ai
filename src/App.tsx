import { useState, useMemo, useCallback, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import html2canvas from 'html2canvas';
import { 
  BookOpen, 
  Upload, 
  Search, 
  Loader2, 
  FileText, 
  ChevronRight, 
  Book as BookIcon,
  X,
  Plus,
  Download,
  MessageSquare,
  Send,
  PieChart,
  BarChart3,
  TrendingUp,
  Sparkles,
  Edit3,
  Copy,
  Check,
  GripVertical,
  Presentation
} from 'lucide-react';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Footer, PageNumber, SectionType, ImageRun } from 'docx';
import pptxgen from "pptxgenjs";
import { saveAs } from 'file-saver';
import { ResponsiveContainer, PieChart as RePieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip } from 'recharts';
import { GoogleGenAI } from "@google/genai";
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragEndEvent,
} from '@dnd-kit/core';
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { cn } from './lib/utils';
import { extractText } from './lib/extractor';
import { generateTopicsFromText } from './lib/gemini';
import { BookTopic, SourceFile } from './types';

export default function App() {
  const [topics, setTopics] = useState<BookTopic[]>([]);
  const [sourceFiles, setSourceFiles] = useState<SourceFile[]>([]);
  const [selectedTopicId, setSelectedTopicId] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [isUploadOpen, setIsUploadOpen] = useState(true);
  const [isChatOpen, setIsChatOpen] = useState(false);
  const [chatMessage, setChatMessage] = useState('');
  const [chatHistory, setChatHistory] = useState<{ role: 'user' | 'assistant', content: string }[]>([]);
  const [isChatLoading, setIsChatLoading] = useState(false);
  const [editingTopicId, setEditingTopicId] = useState<string | null>(null);
  const [editTitle, setEditTitle] = useState('');
  const [coverImageUrl, setCoverImageUrl] = useState<string>('https://images.unsplash.com/photo-1544947950-fa07a98d237f?auto=format&fit=crop&q=80&w=800');

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 8,
      },
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  const sentimentColors: Record<string, string> = {
    analytical: 'bg-vibrant-blue',
    creative: 'bg-vibrant-purple',
    technical: 'bg-vibrant-green',
    narrative: 'bg-vibrant-pink',
  };

  const sentimentTextColors: Record<string, string> = {
    analytical: 'text-vibrant-blue',
    creative: 'text-vibrant-purple',
    technical: 'text-vibrant-green',
    narrative: 'text-vibrant-pink',
  };

  const userEmail = "vivek.impdoc@gmail.com"; // User identifier from metadata

  const filteredTopics = useMemo(() => {
    if (!searchQuery) return topics;
    const query = searchQuery.toLowerCase();
    return topics.filter(t => 
      t.title.toLowerCase().includes(query) || 
      t.summary.toLowerCase().includes(query) ||
      t.content.toLowerCase().includes(query) ||
      t.tags.some(tag => tag.toLowerCase().includes(query))
    );
  }, [topics, searchQuery]);

  const selectedTopic = useMemo(() => 
    topics.find(t => t.id === selectedTopicId), 
    [topics, selectedTopicId]
  );

  const handleFileUpload = useCallback(async (files: FileList | null) => {
    if (!files || files.length === 0) return;
    
    setIsProcessing(true);
    setIsUploadOpen(false);

    try {
      const newTopics: BookTopic[] = [];
      const newSourceFiles: SourceFile[] = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const text = await extractText(file);
        const generated = await generateTopicsFromText(file.name, text);
        newTopics.push(...generated);
        
        newSourceFiles.push({
          name: file.name,
          size: file.size,
          type: file.type,
          lastModified: file.lastModified,
          topicCount: generated.length
        });
      }

      setSourceFiles(prev => [...prev, ...newSourceFiles]);
      setTopics(prev => [...prev, ...newTopics]);
      if (newTopics.length > 0 && !selectedTopicId) {
        setSelectedTopicId(newTopics[0].id);
      }
    } catch (error) {
      console.error("Scale error:", error);
      alert("Error processing files. Please ensure they are valid PDF or DOCX files.");
    } finally {
      setIsProcessing(false);
    }
  }, [selectedTopicId]);

  const fetchImageAsBuffer = async (url: string): Promise<ArrayBuffer> => {
    const response = await fetch(url);
    return await response.arrayBuffer();
  };

  const exportToDocx = useCallback(async () => {
    if (topics.length === 0) return;

    let coverImageRun: ImageRun | null = null;
    try {
      const imgBuffer = await fetchImageAsBuffer(coverImageUrl);
      coverImageRun = new ImageRun({
        data: imgBuffer,
        transformation: {
          width: 500,
          height: 667,
        },
      } as any); // Casting as any to bypass union type resolution issues in some docx versions
    } catch (e) {
      console.error("Failed to fetch cover image for DOCX", e);
    }

    const sentimentHexColors: Record<string, string> = {
      analytical: '1A1A1A',
      skeptical: '7A7870',
      optimistic: 'F59E0B',
      narrative: 'EC4899',
    };

    const sections = [
      // Cover Page Section
      {
        properties: { type: SectionType.NEXT_PAGE },
        children: [
          ...(coverImageRun ? [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [coverImageRun],
              spacing: { before: 400, after: 800 }
            })
          ] : []),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400, after: 400 },
            children: [
              new TextRun({
                text: "ARCHIVE SYNTHESIS SYSTEM",
                size: 20,
                bold: true,
                color: "7A7870",
                characterSpacing: 40,
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 1200 },
            children: [
              new TextRun({
                text: "Scriba AI",
                size: 160,
                font: "Georgia",
                italics: true,
                color: "1A1A1A",
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: "________________________________________________",
                color: "1A1A1A",
              })
            ],
            spacing: { after: 800 }
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 600 },
            children: [
              new TextRun({
                text: "COMPLETE INTELLIGENCE ARCHIVES",
                size: 36,
                bold: true,
                color: "1A1A1A",
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 4000 },
            children: [
              new TextRun({
                text: `Prepared for: ${userEmail}`,
                size: 22,
                italics: true,
                color: "7A7870",
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `Synthesized on ${new Date().toLocaleDateString()}`,
                size: 18,
                color: "999999",
              })
            ]
          })
        ],
      },
      // Content Sections
      ...topics.map((topic, idx) => ({
        properties: { type: SectionType.NEXT_PAGE },
        children: [
          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 400, after: 200 },
            children: [
              new TextRun({
                text: `SECTION 0${idx + 1}`,
                size: 24,
                bold: true,
                color: "7A7870",
                characterSpacing: 20,
              })
            ]
          }),
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 1000 },
            children: [
              new TextRun({
                text: topic.title,
                size: 84,
                font: "Georgia",
                italics: true,
                color: "1A1A1A",
              })
            ]
          }),
          new Paragraph({
            spacing: { after: 800 },
            children: [
              new TextRun({
                text: `TONE: ${topic.sentiment?.toUpperCase() || 'NARRATIVE'} | TAGS: ${topic.tags.join(', ')}`,
                size: 20,
                color: sentimentHexColors[topic.sentiment || 'narrative'],
                bold: true,
              })
            ]
          }),
          new Paragraph({
            spacing: { after: 1400 },
            children: [
              new TextRun({
                text: topic.summary,
                size: 28,
                italics: true,
                color: "444444",
              })
            ],
            indent: { left: 720 },
          }),
          ...topic.content.split('\n').filter(p => p.trim()).map(para => 
            new Paragraph({
              spacing: { after: 400 },
              alignment: AlignmentType.JUSTIFIED,
              indent: { firstLine: 432 },
              children: [
                new TextRun({
                  text: para,
                  size: 24,
                  font: "Georgia",
                })
              ]
            })
          ),
          // Conceptual Flow Section in DOCX
          ...(topic.processSteps && topic.processSteps.length > 0 ? [
            new Paragraph({
              spacing: { before: 1200, after: 600 },
              children: [
                new TextRun({
                  text: "CONCEPTUAL METHODOLOGY",
                  size: 22,
                  bold: true,
                  color: "7A7870",
                  characterSpacing: 10,
                })
              ]
            }),
            ...topic.processSteps.flatMap((step, sIdx) => [
              new Paragraph({
                spacing: { after: 150 },
                children: [
                  new TextRun({
                    text: `${sIdx + 1}. ${step.title.toUpperCase()}`,
                    bold: true,
                    size: 24,
                    color: sentimentHexColors[topic.sentiment || 'narrative'],
                  })
                ]
              }),
              new Paragraph({
                spacing: { after: 300 },
                children: [
                  new TextRun({
                    text: step.description,
                    size: 22,
                    color: "555555",
                  })
                ],
                indent: { left: 432 }
              })
            ])
          ] : []),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 1200 },
            children: [
              new TextRun({
                text: `–– Folio 0${idx + 1} ––`,
                size: 18,
                italics: true,
                color: "7A7870",
              })
            ]
          }),
        ],
      })),
    ];

    const doc = new Document({
      background: {
        color: "F8F7F2",
      },
      sections: sections as any,
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `Scriba_Full_Archive_${new Date().toISOString().split('T')[0]}.docx`);
  }, [topics, userEmail]);

  const exportToPptx = useCallback(() => {
    if (topics.length === 0) return;

    const pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';
    pres.author = 'Scriba AI';
    pres.company = 'TopicBook Synthesizer';
    pres.title = 'Scriba AI Synthesized Archive';

    const bgHex = "F8F7F2";
    const accentHex = "7A7870";
    const textHex = "1A1A1A";

    const sentimentHexColors: Record<string, string> = {
      analytical: '1A1A1A',
      skeptical: '7A7870',
      optimistic: 'F59E0B',
      narrative: 'EC4899',
    };

    // 1. Cover Slide
    const cover = pres.addSlide();
    cover.background = { color: bgHex };
    
    // Add cover image as background or large centered image
    cover.addImage({
      path: coverImageUrl,
      x: 0, y: 0, w: '100%', h: '100%',
      sizing: { type: 'cover', w: 10, h: 5.625 } 
    });

    // Semi-transparent overlay for text readability
    cover.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: '100%', h: '100%',
      fill: { color: "000000", transparency: 60 }
    });
    
    cover.addText("ARCHIVE SYNTHESIS SYSTEM", {
      x: 0, y: 1.5, w: '100%',
      align: 'center',
      fontSize: 14,
      fontFace: 'Arial',
      bold: true,
      color: "FFFFFF",
      charSpacing: 4
    });

    cover.addText("Scriba AI", {
      x: 0, y: 2.2, w: '100%',
      align: 'center',
      fontSize: 72,
      fontFace: 'Georgia',
      italic: true,
      color: "FFFFFF"
    });

    cover.addShape(pres.ShapeType.line, {
      x: 3.5, y: 3.5, w: 3, h: 0,
      line: { color: "FFFFFF", width: 1 }
    });

    cover.addText("COMPLETE INTELLIGENCE ARCHIVES", {
      x: 0, y: 4.0, w: '100%',
      align: 'center',
      fontSize: 24,
      bold: true,
      color: "FFFFFF"
    });

    cover.addText(`Prepared for: ${userEmail}`, {
      x: 0, y: 4.8, w: '100%',
      align: 'center',
      fontSize: 14,
      italic: true,
      color: "CCCCCC"
    });

    cover.addText(`Synthesized on ${new Date().toLocaleDateString()}`, {
      x: 0, y: 5.2, w: '100%',
      align: 'center',
      fontSize: 12,
      color: "AAAAAA"
    });

    // 2. Content Slides
    topics.forEach((topic, idx) => {
      const slide = pres.addSlide();
      slide.background = { color: bgHex };

      // Header info
      slide.addText(`SECTION 0${idx + 1}`, {
        x: 0.5, y: 0.4, w: 2,
        fontSize: 14,
        bold: true,
        fontFace: 'Arial',
        color: accentHex,
        charSpacing: 2
      });

      // Title
      slide.addText(topic.title, {
        x: 0.5, y: 0.8, w: 9,
        fontSize: 48,
        fontFace: 'Georgia',
        italic: true,
        color: textHex
      });

      // Tone/Tags
      slide.addText(`TONE: ${topic.sentiment?.toUpperCase() || 'NARRATIVE'} | TAGS: ${topic.tags.join(', ')}`, {
        x: 0.5, y: 1.6, w: 9,
        fontSize: 11,
        fontFace: 'Arial',
        bold: true,
        color: sentimentHexColors[topic.sentiment || 'narrative']
      });

      // Summary
      slide.addText(topic.summary, {
        x: 1.0, y: 2.2, w: 8,
        fontSize: 18,
        fontFace: 'Georgia',
        italic: true,
        color: "444444"
      });

      // Split content into two columns if needed or just a clean block
      const paragraphs = topic.content.split('\n').filter(p => p.trim()).slice(0, 3);
      if (paragraphs.length > 0) {
        slide.addText(paragraphs.join('\n\n'), {
          x: 0.5, y: 3.2, w: 9, h: 2,
          fontSize: 16,
          fontFace: 'Georgia',
          color: textHex,
          align: 'justify'
        });
      }

      // Conceptual Flow if exists
      if (topic.processSteps && topic.processSteps.length > 0) {
        slide.addText("CONCEPTUAL METHODOLOGY", {
          x: 0.5, y: 5.1, w: 9,
          fontSize: 10,
          bold: true,
          fontFace: 'Arial',
          color: accentHex,
          charSpacing: 1
        });

        const stepW = 9 / Math.min(topic.processSteps.length, 4);
        topic.processSteps.slice(0, 4).forEach((step, sIdx) => {
          const sX = 0.5 + (sIdx * stepW);
          slide.addShape(pres.ShapeType.rect, {
            x: sX, y: 5.4, w: stepW - 0.2, h: 1.4,
            fill: { color: "FFFFFF" },
            line: { color: accentHex, width: 0.5 }
          });
          slide.addText(`${sIdx + 1}. ${step.title.toUpperCase()}`, {
            x: sX + 0.1, y: 5.5, w: stepW - 0.4,
            fontSize: 10,
            bold: true,
            fontFace: 'Arial',
            color: sentimentHexColors[topic.sentiment || 'narrative']
          });
          slide.addText(step.description.substring(0, 80) + '...', {
            x: sX + 0.1, y: 5.8, w: stepW - 0.4,
            fontSize: 9,
            fontFace: 'Georgia',
            color: "555555"
          });
        });
      }

      // Footer
      slide.addText(`–– Folio 0${idx + 1} ––`, {
        x: 0, y: 7.1, w: '100%',
        align: 'center',
        fontSize: 12,
        fontFace: 'Georgia',
        italic: true,
        color: accentHex
      });
    });

    pres.writeFile({ fileName: `Scriba_Full_Archive_${new Date().toISOString().split('T')[0]}.pptx` });
  }, [topics, userEmail]);

  const exportToImage = useCallback(async () => {
    const element = document.getElementById('printable-book');
    if (!element) return;

    const originalStyle = element.getAttribute('style') || '';
    const originalClass = element.className;

    // Force A4 dimensions for high-quality capture
    element.className = "block bg-editorial-bg text-editorial-ink w-[210mm]"; 
    element.style.position = 'fixed';
    element.style.left = '0';
    element.style.top = '0';
    element.style.zIndex = '-9999';
    element.style.visibility = 'visible';
    element.style.printColorAdjust = 'exact';
    (element.style as any).WebkitPrintColorAdjust = 'exact';

    try {
      const canvas = await html2canvas(element, {
        scale: 3, // Even higher quality for print (300 DPI equivalent)
        useCORS: true,
        backgroundColor: '#F8F7F2',
        windowWidth: 794, 
        onclone: (doc) => {
          const el = doc.getElementById('printable-book');
          if (el) {
            el.style.display = 'block';
            el.style.visibility = 'visible';
          }
        }
      });
      
      const link = document.createElement('a');
      link.download = `Scriba_HighRes_Folio_${new Date().toISOString().split('T')[0]}.png`;
      link.href = canvas.toDataURL('image/png', 1.0);
      link.click();
    } catch (error) {
      console.error("Image export failed:", error);
    } finally {
      element.className = originalClass;
      element.style.cssText = originalStyle;
    }
  }, []);

  const exportAll = useCallback(() => {
    // Sequence exports with brief delays to allow browser to handle multiple downloads
    exportToDocx();
    
    setTimeout(() => {
      exportToPptx();
    }, 400);

    setTimeout(() => {
      exportToImage();
    }, 1200);

    setTimeout(() => {
      window.print();
    }, 2000);
  }, [exportToDocx, exportToImage, exportToPptx]);

  const handleChatSend = async (overrideMsg?: string) => {
    const msgToSend = overrideMsg || chatMessage.trim();
    if (!msgToSend || topics.length === 0) return;
    
    if (!overrideMsg) {
      setChatHistory(prev => [...prev, { role: 'user', content: msgToSend }]);
      setChatMessage('');
    }
    setIsChatLoading(true);

    try {
      const context = topics.map(t => {
        let text = `Topic: ${t.title}\nSummary: ${t.summary}\n`;
        if (t.processSteps && t.processSteps.length > 0) {
          text += `Key Process Steps:\n${t.processSteps.map(s => `- ${s.title}: ${s.description}`).join('\n')}\n`;
        }
        text += `Content Analysis: ${t.content.substring(0, 800)}...`;
        return text;
      }).join('\n\n---\n\n');
      
      const genAI = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
      
      const systemPrompt = `You are "Scriba AI Oracle", a high-level research assistant. 
      The user has synthesized an archive with ${topics.length} chapters.
      
      ARCHIVE CONTEXT:
      ${context}
      
      MISSION:
      - Use the provided context to answer accurately.
      - If asked for a summary, provide a synthesis of all chapters.
      - If the user just opened the chat, give a warm, intellectual greeting summarizing what their archive covers.
      - Be academic yet accessible.
      
      User Message: ${msgToSend}`;

      const result = await genAI.models.generateContent({
        model: "gemini-1.5-flash",
        contents: systemPrompt
      });
      const responseText = result.text || "I apologize, but I encountered an error in my neural processing.";
      
      setChatHistory(prev => [...prev, { role: 'assistant', content: responseText }]);
    } catch (e) {
      console.error(e);
      setChatHistory(prev => [...prev, { role: 'assistant', content: "Mafi chahta hoon, main abhi process nahi kar pa raha hoon. Technical issue detected." }]);
    } finally {
      setIsChatLoading(false);
    }
  };

  // Initial greeting when chat is opened
  useEffect(() => {
    if (isChatOpen && chatHistory.length === 0 && topics.length > 0) {
      handleChatSend("Hello! Please introduce yourself and provide a brief overview of my synthesized archive.");
    }
  }, [isChatOpen, topics.length]);

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;

    if (over && active.id !== over.id) {
      setTopics((items) => {
        const oldIndex = items.findIndex((t) => t.id === active.id);
        const newIndex = items.findIndex((t) => t.id === over.id);
        return arrayMove(items, oldIndex, newIndex);
      });
    }
  };

  const handleCopyTopic = (topic: BookTopic) => {
    const textToCopy = `# ${topic.title}\n\n${topic.summary}\n\n${topic.content}`;
    navigator.clipboard.writeText(textToCopy).then(() => {
      alert('Topic content copied to clipboard');
    });
  };

  const startEditing = (topic: BookTopic) => {
    setEditingTopicId(topic.id);
    setEditTitle(topic.title);
  };

  const saveEdit = (topicId: string) => {
    setTopics(prev => prev.map(t => t.id === topicId ? { ...t, title: editTitle } : t));
    setEditingTopicId(null);
  };

  return (
    <div className="min-h-screen bg-editorial-bg text-editorial-ink font-sans flex overflow-hidden selection:bg-[#E4D8C1]">
      {/* Left Sidebar: Asset Management */}
      <aside className="w-72 border-r border-editorial-border bg-editorial-sidebar flex flex-col h-screen overflow-hidden">
        <div className="p-8 border-b border-editorial-border">
          <h1 className="text-2xl font-serif italic tracking-tight mb-1">Scriba AI</h1>
          <p className="text-[10px] uppercase tracking-widest text-editorial-muted font-bold">TopicBook Synthesizer</p>
        </div>
        
        <div className="flex-1 p-6 space-y-8 overflow-y-auto custom-scrollbar">
          {/* Sources Section */}
          <section>
            <h3 className="text-[11px] uppercase tracking-widest text-editorial-muted mb-4 font-bold">Sources ({sourceFiles.length})</h3>
            <div className="space-y-3">
              {sourceFiles.length === 0 && (
                <p className="text-[10px] italic text-editorial-muted opacity-60">No files uploaded</p>
              )}
              {sourceFiles.map(file => (
                <div 
                  key={file.name} 
                  className="flex items-center text-xs group cursor-pointer text-editorial-ink/80 hover:text-editorial-ink transition-all"
                  onClick={() => setSelectedTopicId(null)}
                >
                  <div className="w-2 h-2 rounded-full bg-editorial-ink mr-3 shrink-0"></div>
                  <span className="flex-1 truncate">{file.name}</span>
                  <span className="text-[9px] text-[#A5A296] opacity-0 group-hover:opacity-100 transition-opacity bg-editorial-accent-bg px-1 rounded uppercase font-bold">Details</span>
                </div>
              ))}
              <button 
                onClick={() => setIsUploadOpen(true)}
                className="mt-4 w-full py-2 border border-dashed border-[#A5A296] text-[10px] uppercase tracking-widest hover:bg-[#EAE8E0] transition-colors font-bold"
              >
                + Upload New File
              </button>
            </div>
          </section>

          {/* Chapters Section */}
          <section>
            <h3 className="text-[11px] uppercase tracking-widest text-editorial-muted mb-4 font-bold">Generated Topics</h3>
            <nav className="space-y-1">
              {filteredTopics.length === 0 && !isProcessing && (
                <p className="text-[10px] italic text-editorial-muted opacity-60">No topics analyzed yet</p>
              )}
              {isProcessing && (
                <div className="flex items-center gap-2 text-xs italic text-editorial-muted">
                  <Loader2 className="w-3 h-3 animate-spin" />
                  Synthesizing...
                </div>
              )}
              <DndContext
                sensors={sensors}
                collisionDetection={closestCenter}
                onDragEnd={handleDragEnd}
              >
                <SortableContext
                  items={filteredTopics.map((t) => t.id)}
                  strategy={verticalListSortingStrategy}
                >
                  {filteredTopics.map((topic, idx) => (
                    <SortableTopicItem
                      key={topic.id}
                      topic={topic}
                      idx={idx}
                      isSelected={selectedTopicId === topic.id}
                      onSelect={() => setSelectedTopicId(topic.id)}
                      isEditing={editingTopicId === topic.id}
                      editTitle={editTitle}
                      onEditChange={setEditTitle}
                      onStartEdit={() => startEditing(topic)}
                      onSaveEdit={() => saveEdit(topic.id)}
                      onCopy={() => handleCopyTopic(topic)}
                      sentimentColors={sentimentColors}
                    />
                  ))}
                </SortableContext>
              </DndContext>
            </nav>
          </section>
        </div>

        {/* Search Panel */}
        <div className="p-6 bg-editorial-accent-bg border-t border-editorial-border">
          <div className="relative group">
            <input 
              type="text" 
              placeholder="Search chapters..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full bg-white border border-editorial-border px-3 py-2 text-xs focus:outline-none italic placeholder:text-editorial-muted/40"
            />
            <div className="absolute right-3 top-2.5 opacity-40">
              <Search className="w-3.5 h-3.5" />
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content: The Book Preview */}
      <main className="flex-1 flex flex-col h-screen overflow-hidden">
        {/* Header/Toolbar */}
        <header className="h-16 border-b border-editorial-border flex items-center justify-between px-10 bg-editorial-bg shrink-0">
          <div className="flex space-x-8">
            <button className="text-[10px] uppercase tracking-widest font-bold border-b-2 border-editorial-ink py-1">Layout View</button>
            <button className="text-[10px] uppercase tracking-widest text-editorial-muted hover:text-editorial-ink transition-colors py-1">Typography</button>
            <button className="text-[10px] uppercase tracking-widest text-editorial-muted hover:text-editorial-ink transition-colors py-1">Insights</button>
          </div>
          <div className="flex items-center gap-4">
            <button 
              className="bg-black text-white px-6 py-2 text-[10px] uppercase tracking-widest hover:opacity-80 transition-opacity font-bold disabled:opacity-20 disabled:cursor-not-allowed flex items-center gap-2"
              onClick={exportAll}
              disabled={topics.length === 0}
            >
              <Download className="w-3 h-3" />
              Generate Book (Word + PPT + PDF + Image)
            </button>
            <div className="w-[1px] h-6 bg-editorial-border mx-2"></div>
            <button 
              className="text-editorial-muted hover:text-editorial-ink text-[10px] uppercase tracking-widest font-bold transition-colors"
              onClick={exportToDocx}
              disabled={topics.length === 0}
            >
              MS Word
            </button>
            <button 
              className="text-editorial-muted hover:text-editorial-ink text-[10px] uppercase tracking-widest font-bold transition-colors"
              onClick={exportToPptx}
              disabled={topics.length === 0}
            >
              PowerPoint
            </button>
            <button 
              className="text-editorial-muted hover:text-editorial-ink text-[10px] uppercase tracking-widest font-bold transition-colors"
              onClick={() => window.print()}
              disabled={topics.length === 0}
            >
              PDF Archive
            </button>
            <button 
              className="text-editorial-muted hover:text-editorial-ink text-[10px] uppercase tracking-widest font-bold transition-colors"
              onClick={exportToImage}
              disabled={topics.length === 0}
            >
              Image Record
            </button>
          </div>
        </header>

        {/* Page Preview Area */}
        <div className="flex-1 overflow-y-auto p-12 flex justify-center items-start bg-editorial-accent-bg custom-scrollbar">
          <AnimatePresence mode="wait">
            {selectedTopic ? (
              <motion.article
                key={selectedTopic.id}
                initial={{ opacity: 0, y: 30 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -30 }}
                transition={{ duration: 0.5, ease: [0.19, 1, 0.22, 1] }}
                className="w-full max-w-[580px] min-h-[740px] bg-white shadow-[0_30px_60px_-15px_rgba(0,0,0,0.1)] p-16 flex flex-col relative mb-12"
              >
                {/* Page Decoration */}
                <div className="absolute top-10 right-10 font-serif italic text-sm text-editorial-muted">
                  <span className="text-[10px] uppercase tracking-widest font-black mr-2 not-italic opacity-40">Folio</span>
                  {topics.findIndex(t => t.id === selectedTopic.id) + 1}
                </div>

                <div className="mb-12">
                  <span className="text-[9px] uppercase tracking-[0.4em] text-[#A5A296] block mb-3 font-bold">
                    Chapter Selection
                  </span>
                  <h2 className="text-5xl font-serif italic leading-[1.1] mb-8 text-editorial-ink">
                    {selectedTopic.title}
                  </h2>
                  <div className="flex gap-4 items-center mb-8">
                    <div className="w-24 h-[1px] bg-editorial-ink"></div>
                    {selectedTopic.sentiment && (
                      <span className={cn(
                        "text-[9px] uppercase tracking-widest font-black px-2 py-0.5 rounded border border-current",
                        sentimentTextColors[selectedTopic.sentiment]
                      )}>
                        {selectedTopic.sentiment} Tone
                      </span>
                    )}
                    <span className="text-[9px] uppercase tracking-widest font-black text-editorial-muted">
                      ~{selectedTopic.wordCount} Words
                    </span>
                  </div>
                </div>

                <div className="prose prose-stone max-w-none book-content space-y-6 flex-1">
                  {selectedTopic.content.split('\n').map((para, i) => (
                    para.trim() && (
                      <p key={i} className="font-serif text-[1.1rem] leading-relaxed text-justify text-editorial-ink/90">
                        {para}
                      </p>
                    )
                  ))}

                  {/* Visual Process Flow / Chart */}
                  {selectedTopic.processSteps && selectedTopic.processSteps.length > 0 && (
                    <div className="my-12 p-8 bg-editorial-sidebar/50 border border-editorial-border rounded-xl">
                      <div className="flex items-center gap-2 mb-8">
                        <TrendingUp className="w-4 h-4 text-vibrant-purple" />
                        <h4 className="text-[10px] uppercase tracking-widest font-black text-editorial-muted">Conceptual Flow</h4>
                      </div>
                      <div className="space-y-6 relative">
                        {selectedTopic.processSteps.map((step, idx) => (
                          <div key={idx} className="relative pl-12">
                            {/* Number & Line */}
                            <div className="absolute left-0 top-0 bottom-0 w-8 flex flex-col items-center">
                              <div className={cn(
                                "w-6 h-6 rounded-full flex items-center justify-center text-white text-[10px] font-black z-10",
                                sentimentColors[selectedTopic.sentiment || 'analytical']
                              )}>
                                {idx + 1}
                              </div>
                              {idx < (selectedTopic.processSteps?.length || 0) - 1 && (
                                <div className="w-[1px] flex-1 bg-editorial-border my-2 border-dashed border"></div>
                              )}
                            </div>
                            <div className="pb-4">
                              <h5 className="font-serif italic text-lg leading-none mb-1">{step.title}</h5>
                              <p className="text-xs text-editorial-muted leading-relaxed">{step.description}</p>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Visual Illustration Section */}
                  <div className="my-12">
                    <div className={cn(
                      "relative aspect-[16/9] border border-editorial-border rounded-xl overflow-hidden group shadow-inner",
                      selectedTopic.sentiment === 'analytical' ? 'bg-gradient-to-br from-blue-50 to-blue-100' :
                      selectedTopic.sentiment === 'creative' ? 'bg-gradient-to-br from-purple-50 to-purple-100' :
                      selectedTopic.sentiment === 'technical' ? 'bg-gradient-to-br from-green-50 to-green-100' :
                      'bg-gradient-to-br from-pink-50 to-pink-100'
                    )}>
                      {/* Geometric Decorative Pattern */}
                      <div className="absolute inset-0 opacity-[0.05] pointer-events-none">
                        <div className="absolute top-0 left-0 w-full h-full" style={{ backgroundImage: 'radial-gradient(circle at 2px 2px, black 1px, transparent 0)', backgroundSize: '24px 24px' }}></div>
                      </div>

                      <div className="absolute inset-0 flex flex-col items-center justify-center text-center p-12 transition-all group-hover:scale-105 duration-700">
                        <div className={cn(
                          "w-20 h-20 rounded-2xl flex items-center justify-center mb-6 shadow-xl",
                          sentimentColors[selectedTopic.sentiment || 'analytical'],
                          "text-white"
                        )}>
                          {selectedTopic.sentiment === 'analytical' ? <TrendingUp className="w-10 h-10" /> :
                           selectedTopic.sentiment === 'creative' ? <Sparkles className="w-10 h-10" /> :
                           selectedTopic.sentiment === 'technical' ? <BarChart3 className="w-10 h-10" /> :
                           <BookIcon className="w-10 h-10" />}
                        </div>
                        <h4 className="text-sm font-serif italic text-editorial-ink/60 max-w-xs">{selectedTopic.title}</h4>
                        <div className="mt-4 flex gap-2">
                          {selectedTopic.tags.slice(0, 3).map(tag => (
                            <span key={tag} className="text-[8px] uppercase tracking-widest font-black text-editorial-muted bg-white/50 px-2 py-0.5 rounded">#{tag}</span>
                          ))}
                        </div>
                      </div>
                      
                      <div className="absolute bottom-6 right-6 opacity-0 group-hover:opacity-100 transition-opacity">
                        <div className="bg-black/90 backdrop-blur-sm text-white px-3 py-1 text-[8px] uppercase tracking-[0.2em] font-black rounded-full shadow-lg">
                          Visual synthesis v1.0
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <footer className="mt-16 pt-8 border-t border-editorial-sidebar flex justify-between items-end">
                  <div className="space-y-1">
                    <p className="text-[9px] uppercase tracking-widest text-editorial-muted font-bold">Synthesized Content</p>
                    <p className="text-[10px] italic text-editorial-muted truncate max-w-[200px]">{selectedTopic.sourceFile}</p>
                  </div>
                  <div className="text-right">
                    <p className="text-[9px] uppercase tracking-widest text-editorial-muted font-bold">SCRIBA AI / 2026</p>
                    <p className="text-[9px] text-editorial-muted/40 uppercase tracking-widest mt-1">Ref: {selectedTopic.id}</p>
                  </div>
                </footer>
              </motion.article>
            ) : sourceFiles.length > 0 ? (
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="w-full max-w-4xl mx-auto py-16 px-12"
              >
                <div className="mb-12">
                  <p className="text-[10px] uppercase tracking-[0.4em] text-editorial-muted font-bold mb-3">Library Management</p>
                  <h2 className="text-4xl font-serif italic text-editorial-ink">Synthesized Archives</h2>
                </div>

                <div className="bg-white border border-editorial-border rounded-xl overflow-hidden shadow-sm">
                  <table className="w-full text-left text-sm border-collapse">
                    <thead className="bg-editorial-sidebar border-b border-editorial-border">
                      <tr>
                        <th className="px-6 py-4 text-[10px] uppercase tracking-widest text-editorial-muted font-bold">File Name</th>
                        <th className="px-6 py-4 text-[10px] uppercase tracking-widest text-editorial-muted font-bold text-center">Type</th>
                        <th className="px-6 py-4 text-[10px] uppercase tracking-widest text-editorial-muted font-bold text-center">Size</th>
                        <th className="px-6 py-4 text-[10px] uppercase tracking-widest text-editorial-muted font-bold text-center">Topics</th>
                        <th className="px-6 py-4 text-[10px] uppercase tracking-widest text-editorial-muted font-bold text-right">Added</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-editorial-border">
                      {sourceFiles.map((file) => (
                        <tr key={file.name} className="hover:bg-editorial-bg transition-colors">
                          <td className="px-6 py-4 font-medium flex items-center gap-3">
                            <FileText className="w-4 h-4 text-editorial-muted" />
                            {file.name}
                          </td>
                          <td className="px-6 py-4 text-center text-xs opacity-60">
                            {file.type.split('/')[1]?.toUpperCase() || 'DOCX'}
                          </td>
                          <td className="px-6 py-4 text-center text-xs opacity-60">
                            {(file.size / 1024).toFixed(1)} KB
                          </td>
                          <td className="px-6 py-4 text-center">
                            <span className="bg-editorial-ink text-white px-2 py-0.5 rounded-full text-[10px] font-bold">
                              {file.topicCount}
                            </span>
                          </td>
                          <td className="px-6 py-4 text-right text-[10px] opacity-40 italic">
                            {new Date(file.lastModified).toLocaleDateString()}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="mt-12 grid grid-cols-3 gap-8">
                  <div className="p-6 bg-white border border-editorial-border rounded-xl">
                    <p className="text-[10px] uppercase tracking-widest text-editorial-muted font-bold mb-2">Total Topics</p>
                    <p className="text-3xl font-serif italic leading-none">{topics.length}</p>
                  </div>
                  <div className="p-6 bg-white border border-editorial-border rounded-xl">
                    <p className="text-[10px] uppercase tracking-widest text-editorial-muted font-bold mb-2">Storage Used</p>
                    <p className="text-3xl font-serif italic leading-none">
                      {(sourceFiles.reduce((acc, f) => acc + f.size, 0) / 1024).toFixed(1)} KB
                    </p>
                  </div>
                  <div className="p-6 bg-white border border-editorial-border rounded-xl">
                    <p className="text-[10px] uppercase tracking-widest text-editorial-muted font-bold mb-2">Avg. Depth</p>
                    <p className="text-3xl font-serif italic leading-none">
                      {(topics.length / sourceFiles.length).toFixed(1)} <span className="text-xs italic opacity-40">T/F</span>
                    </p>
                  </div>
                </div>

                {/* Vibrant Data Charts */}
                <div className="mt-12 grid grid-cols-2 gap-8">
                  <div className="p-8 bg-white border border-editorial-border rounded-2xl shadow-sm h-[300px]">
                    <div className="flex items-center gap-2 mb-6">
                      <PieChart className="w-4 h-4 text-vibrant-purple" />
                      <p className="text-[10px] uppercase tracking-widest text-editorial-muted font-bold">Sentiment Distribution</p>
                    </div>
                    <ResponsiveContainer width="100%" height="80%">
                      <RePieChart>
                        <Pie
                          data={Object.keys(sentimentColors).map(s => ({
                            name: s,
                            value: topics.filter(t => t.sentiment === s).length
                          })).filter(d => d.value > 0)}
                          innerRadius={60}
                          outerRadius={80}
                          paddingAngle={5}
                          dataKey="value"
                        >
                          {['analytical', 'creative', 'technical', 'narrative'].map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={
                              entry === 'analytical' ? '#3B82F6' : 
                              entry === 'creative' ? '#8B5CF6' : 
                              entry === 'technical' ? '#10B981' : '#EC4899'
                            } />
                          ))}
                        </Pie>
                        <Tooltip />
                      </RePieChart>
                    </ResponsiveContainer>
                  </div>

                  <div className="p-8 bg-white border border-editorial-border rounded-2xl shadow-sm h-[300px]">
                    <div className="flex items-center gap-2 mb-6 text-vibrant-blue">
                      <BarChart3 className="w-4 h-4" />
                      <p className="text-[10px] uppercase tracking-widest text-editorial-muted font-bold">Content Density (Words)</p>
                    </div>
                    <ResponsiveContainer width="100%" height="80%">
                      <BarChart data={topics.slice(0, 6)}>
                        <XAxis dataKey="title" hide />
                        <YAxis hide />
                        <Tooltip />
                        <Bar dataKey="wordCount" radius={[4, 4, 0, 0]}>
                          {topics.map((t, index) => (
                            <Cell key={`at-${index}`} fill={
                              t.sentiment === 'analytical' ? '#3B82F6' : 
                              t.sentiment === 'creative' ? '#8B5CF6' : 
                              t.sentiment === 'technical' ? '#10B981' : '#EC4899'
                            } />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </motion.div>
            ) : (
              <div className="h-full flex flex-col items-center justify-center text-center opacity-40 scale-110 mt-20">
                <BookIcon className="w-24 h-24 mb-8 text-editorial-muted" />
                <h2 className="font-serif italic text-3xl mb-2">The Archive is Empty</h2>
                <p className="text-xs uppercase tracking-widest font-bold">Upload documents to begin synthesis</p>
              </div>
            )}
          </AnimatePresence>
        </div>
      </main>

      {/* Right Narrow Panel: Utility */}
      <aside className="w-16 border-l border-editorial-border bg-editorial-bg flex flex-col items-center py-10 space-y-10 shrink-0">
        <button className="p-2.5 hover:bg-editorial-sidebar transition-colors rounded text-editorial-muted hover:text-editorial-ink" title="Page Settings">
          <BookIcon className="w-5 h-5" />
        </button>
        <button className="p-2.5 hover:bg-editorial-sidebar transition-colors rounded text-editorial-muted hover:text-editorial-ink" title="AI Refresh">
          <Loader2 className="w-5 h-5" />
        </button>
        <button 
          className={cn(
            "p-2.5 transition-all rounded",
            isChatOpen ? "bg-vibrant-purple text-white shadow-lg" : "hover:bg-editorial-sidebar text-editorial-muted hover:text-editorial-ink"
          )}
          onClick={() => setIsChatOpen(!isChatOpen)}
          title="AI Discussion"
        >
          <MessageSquare className="w-5 h-5" />
        </button>
        <button className="p-2.5 bg-black text-white transition-colors rounded" title="Smart Search">
          <Search className="w-5 h-5" />
        </button>
        <div className="flex-1"></div>
        <div className="text-[10px] uppercase tracking-widest text-editorial-muted font-black vertical-rl transform rotate-180 mb-6 py-2 px-1">
          Beta v1.2.4
        </div>
      </aside>

      {/* Floating Chat Interface */}
      <AnimatePresence>
        {isChatOpen && (
          <motion.div
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: 20 }}
            className="fixed bottom-24 right-20 w-80 h-[500px] bg-white border border-editorial-border shadow-2xl rounded-2xl flex flex-col overflow-hidden z-40"
          >
            <div className="p-4 border-b border-editorial-border bg-editorial-sidebar flex justify-between items-center">
              <div className="flex items-center gap-2">
                <Sparkles className="w-4 h-4 text-vibrant-purple" />
                <span className="text-xs uppercase tracking-widest font-black">Archive Oracle</span>
              </div>
              <button 
                onClick={() => setIsChatOpen(false)}
                className="opacity-40 hover:opacity-100 transition-opacity"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-4 space-y-4 custom-scrollbar bg-editorial-bg/30">
              {chatHistory.length === 0 && (
                <div className="h-full flex flex-col items-center justify-center text-center opacity-40 px-6">
                  <MessageSquare className="w-8 h-8 mb-4" />
                  <p className="text-xs font-serif italic">Ask me anything about your synthesized chapters.</p>
                </div>
              )}
              {chatHistory.map((msg, i) => (
                <div key={i} className={cn(
                  "max-w-[85%] rounded-2xl p-3 text-sm",
                  msg.role === 'user' 
                    ? "ml-auto bg-editorial-ink text-white rounded-br-none" 
                    : "mr-auto bg-white border border-editorial-border rounded-bl-none text-editorial-ink shadow-sm"
                )}>
                  {msg.content}
                </div>
              ))}
              {isChatLoading && (
                <div className="mr-auto bg-white border border-editorial-border rounded-2xl rounded-bl-none p-3 shadow-sm">
                  <div className="flex gap-1">
                    <div className="w-1.5 h-1.5 rounded-full bg-editorial-muted animate-bounce" />
                    <div className="w-1.5 h-1.5 rounded-full bg-editorial-muted animate-bounce [animation-delay:0.2s]" />
                    <div className="w-1.5 h-1.5 rounded-full bg-editorial-muted animate-bounce [animation-delay:0.4s]" />
                  </div>
                </div>
              )}
            </div>

            <div className="p-4 border-t border-editorial-border bg-white">
              <div className="relative">
                <input 
                  type="text"
                  placeholder="Discuss archives..."
                  value={chatMessage}
                  onChange={(e) => setChatMessage(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleChatSend()}
                  className="w-full bg-editorial-bg border border-editorial-border rounded-xl px-4 py-2 text-xs focus:ring-1 focus:ring-vibrant-purple transition-all pr-10"
                />
                <button 
                  onClick={handleChatSend}
                  disabled={!chatMessage.trim() || isChatLoading}
                  className="absolute right-2 top-1.5 p-1 text-vibrant-purple disabled:opacity-20 translate-y-0.5"
                >
                  <Send className="w-4 h-4" />
                </button>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Hidden container for Printing/PDF Export */}
      <div 
        id="printable-book" 
        className="hidden print:block absolute inset-0 z-[-1] bg-editorial-bg text-editorial-ink"
        style={{ printColorAdjust: 'exact', WebkitPrintColorAdjust: 'exact' } as any}
      >
        {/* PDF Cover Page */}
        <div className="print-page-break flex flex-col items-center justify-center text-center p-20 min-h-screen relative overflow-hidden bg-white">
          <img 
            src={coverImageUrl} 
            alt="Cover" 
            className="absolute inset-0 w-full h-full object-cover opacity-15 grayscale brightness-125"
            referrerPolicy="no-referrer"
          />
          <div className="relative z-10 w-full flex flex-col items-center">
            <div className="mb-20">
              <p className="text-[10px] uppercase tracking-[0.8em] text-editorial-muted mb-6 font-black">Archive Synthesis System</p>
              <h1 className="text-[10rem] font-serif italic tracking-tighter leading-none mb-4 text-editorial-ink">Scriba AI</h1>
            </div>
            <div className="w-32 h-[1px] bg-editorial-ink mb-12"></div>
            <div className="space-y-4">
              <h2 className="text-4xl uppercase tracking-[0.4em] font-light text-editorial-ink mb-8">Intelligence Archives</h2>
              <div className="p-10 border border-editorial-border bg-editorial-bg/30 backdrop-blur-md inline-block">
                <p className="text-[10px] uppercase tracking-widest font-black text-editorial-muted mb-2">Authenticated For:</p>
                <p className="text-lg font-serif italic font-medium">{userEmail}</p>
              </div>
            </div>
            
            <div className="mt-40 grid grid-cols-2 gap-20 w-full max-w-2xl text-left">
              <div>
                <p className="text-[9px] uppercase tracking-[0.2em] font-black border-b border-editorial-border pb-2 mb-4">Metric Analysis</p>
                <div className="space-y-4">
                  <div className="flex justify-between items-end">
                    <span className="text-xs italic font-serif">Total Chapters</span>
                    <span className="text-xl font-serif">{topics.length}</span>
                  </div>
                  <div className="flex justify-between items-end">
                    <span className="text-xs italic font-serif">Total Volume</span>
                    <span className="text-xl font-serif">~{topics.reduce((acc, t) => acc + t.wordCount, 0)} Words</span>
                  </div>
                </div>
              </div>
              <div>
                <p className="text-[9px] uppercase tracking-[0.2em] font-black border-b border-editorial-border pb-2 mb-4">Temporal Data</p>
                <div className="space-y-4 text-xs italic font-serif">
                  <p>Synthesized: {new Date().toLocaleDateString()}</p>
                  <p>Status: Master Archive Record</p>
                </div>
              </div>
            </div>

            <div className="absolute bottom-20 left-20 right-20 flex justify-between text-[9px] uppercase tracking-[0.5em] font-black text-editorial-muted">
              <span>Ver. 1.2.4 Beta</span>
              <span>Ref: {Math.random().toString(36).substring(7).toUpperCase()}</span>
            </div>
          </div>
        </div>

        {topics.map((topic, idx) => (
          <div key={topic.id} className="print-page-break p-12">
            <div className="max-w-[700px] mx-auto min-h-full flex flex-col">
              {/* Header */}
              <div className="flex justify-between items-baseline mb-16 border-b border-editorial-border pb-4">
                <span className="text-[10px] uppercase tracking-[0.5em] font-black text-editorial-muted">
                  Sec. 0{idx + 1}
                </span>
                <span className="text-[9px] uppercase tracking-widest font-black text-editorial-muted italic">
                  Archive Intel
                </span>
              </div>

              <div className="mb-20">
                <div className="flex items-center gap-4 mb-6">
                  <div className={cn("w-12 h-[1px]", sentimentColors[topic.sentiment || 'analytical'])} />
                  <span className={cn(
                    "text-[10px] uppercase tracking-[0.3em] font-black",
                    sentimentTextColors[topic.sentiment || 'analytical']
                  )}>
                    {topic.sentiment} Perspective
                  </span>
                </div>
                <h2 className="text-6xl font-serif italic leading-[1.05] mb-10 text-editorial-ink">
                  {topic.title}
                </h2>
                <p className="text-sm border-l-2 border-editorial-border pl-6 italic text-editorial-muted leading-relaxed max-w-xl">
                  {topic.summary}
                </p>
              </div>

              <div className="book-content prose prose-stone max-w-none space-y-8 flex-1">
                {topic.content.split('\n').map((para, i) => (
                  para.trim() && (
                    <p key={i} className="font-serif text-[1.15rem] leading-relaxed text-justify text-black indent-8">
                      {para}
                    </p>
                  )
                ))}

                {/* Process Steps in Print */}
                {topic.processSteps && topic.processSteps.length > 0 && (
                  <div className="mt-20 pt-12 border-t border-editorial-border">
                    <p className="text-[9px] uppercase tracking-[0.4em] font-black text-editorial-muted mb-10">Conceptual Methodology</p>
                    <div className="grid grid-cols-2 gap-10">
                      {topic.processSteps.map((step, sIdx) => (
                        <div key={sIdx} className="space-y-2">
                          <div className="flex items-center gap-3">
                            <span className={cn(
                              "w-5 h-5 rounded-full flex items-center justify-center text-[10px] text-white font-black",
                              sentimentColors[topic.sentiment || 'analytical']
                            )}>
                              {sIdx + 1}
                            </span>
                            <h5 className="font-serif italic text-lg">{step.title}</h5>
                          </div>
                          <p className="text-xs leading-relaxed text-editorial-muted pl-8">{step.description}</p>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              <footer className="mt-20 pt-8 border-t border-editorial-border flex justify-between items-end italic opacity-80 text-[10px]">
                <div className="space-y-1">
                  <p className="uppercase tracking-[0.3em] font-black not-italic text-[8px] text-editorial-muted">Source Authenticity</p>
                  <p className="max-w-[400px] truncate font-serif text-xs">{topic.sourceFile}</p>
                </div>
                <div className="text-right">
                  <p className="uppercase tracking-[0.3em] font-black not-italic text-[8px] text-editorial-muted mb-2">Folio 0{idx + 1}</p>
                  <div className={cn("inline-block px-3 py-1 text-white font-serif italic text-lg shadow-sm", sentimentColors[topic.sentiment || 'analytical'])}>
                    {idx + 1}
                  </div>
                </div>
              </footer>
            </div>
          </div>
        ))}
      </div>

      {/* Upload Modal */}
      <AnimatePresence>
        {isUploadOpen && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-[#1A1A1A]/80 backdrop-blur-sm z-50 flex items-center justify-center p-6"
          >
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-[#F5F2ED] w-full max-w-xl rounded-3xl overflow-hidden shadow-2xl"
            >
              <div className="p-8 border-b border-[#1A1A1A]/5 flex justify-between items-center">
                <div>
                  <h2 className="font-serif text-2xl font-medium">Add to Library</h2>
                  <p className="text-sm text-[#1A1A1A]/40 mt-1">Upload PDF or DOCX files to organize</p>
                </div>
                <button 
                  onClick={() => setIsUploadOpen(false)}
                  className="p-2 hover:bg-[#1A1A1A]/5 rounded-full transition-colors"
                >
                  <X className="w-5 h-5 text-[#1A1A1A]/40" />
                </button>
              </div>

              <div className="p-8">
                <div 
                  className="border-2 border-dashed border-[#1A1A1A]/10 rounded-2xl p-12 text-center group hover:border-[#1A1A1A]/30 transition-all cursor-pointer relative"
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={(e) => {
                    e.preventDefault();
                    handleFileUpload(e.dataTransfer.files);
                  }}
                >
                  <input 
                    type="file" 
                    multiple 
                    onChange={(e) => handleFileUpload(e.target.files)}
                    className="absolute inset-0 opacity-0 cursor-pointer"
                    accept=".pdf,.docx"
                  />
                  <div className="w-16 h-16 bg-[#1A1A1A]/5 rounded-full flex items-center justify-center mx-auto mb-4 group-hover:scale-110 transition-transform">
                    <Upload className="w-8 h-8 text-[#1A1A1A]/60" />
                  </div>
                  <h3 className="font-serif text-lg font-medium mb-2">Drop your files here</h3>
                  <p className="text-xs text-[#1A1A1A]/40 uppercase tracking-widest font-bold">PDF or DOCX files supported</p>
                </div>

                <div className="mt-8 flex justify-end gap-3">
                  <button 
                    onClick={() => setIsUploadOpen(false)}
                    className="px-6 py-2 text-sm font-medium text-[#1A1A1A]/60 hover:text-[#1A1A1A] transition-colors"
                  >
                    Cancel
                  </button>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

interface SortableTopicItemProps {
  key?: string | number;
  topic: BookTopic;
  idx: number;
  isSelected: boolean;
  onSelect: () => void;
  isEditing: boolean;
  editTitle: string;
  onEditChange: (val: string) => void;
  onStartEdit: () => void;
  onSaveEdit: () => void;
  onCopy: () => void;
  sentimentColors: Record<string, string>;
}

function SortableTopicItem({ 
  topic, 
  idx, 
  isSelected, 
  onSelect, 
  isEditing, 
  editTitle, 
  onEditChange, 
  onStartEdit, 
  onSaveEdit, 
  onCopy,
  sentimentColors
}: SortableTopicItemProps) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging
  } = useSortable({ id: topic.id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    zIndex: isDragging ? 50 : 'auto',
    opacity: isDragging ? 0.5 : 1,
  };

  return (
    <div
      ref={setNodeRef}
      style={style}
      className={cn(
        "group relative flex items-center gap-1 w-full text-left py-2 text-sm font-serif italic border-b transition-all",
        isSelected 
          ? "border-editorial-ink font-bold pl-2 bg-editorial-accent-bg/30" 
          : "border-editorial-border hover:pl-2 text-editorial-ink/70"
      )}
    >
      <div 
        {...attributes} 
        {...listeners}
        className="opacity-0 group-hover:opacity-40 cursor-grab active:cursor-grabbing p-1 -ml-1 hover:opacity-100 transition-opacity"
      >
        <GripVertical className="w-3 h-3" />
      </div>

      {isEditing ? (
        <div className="flex-1 flex items-center gap-2 pr-2">
          <input
            type="text"
            value={editTitle}
            onChange={(e) => onEditChange(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && onSaveEdit()}
            autoFocus
            className="flex-1 bg-white border border-editorial-border px-2 py-0.5 text-xs font-sans not-italic focus:outline-none"
          />
          <button onClick={onSaveEdit} className="text-vibrant-green">
            <Check className="w-3.5 h-3.5" />
          </button>
        </div>
      ) : (
        <div className="flex-1 flex justify-between items-center pr-2 cursor-pointer" onClick={onSelect}>
          <span className="truncate">{String(idx + 1).padStart(2, '0')}. {topic.title}</span>
          <div className="flex items-center gap-2">
            <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
              <button 
                onClick={(e) => { e.stopPropagation(); onCopy(); }} 
                className="p-1 hover:bg-editorial-border rounded transition-colors text-editorial-muted hover:text-editorial-ink"
                title="Copy Topic"
              >
                <Copy className="w-3 h-3" />
              </button>
              <button 
                onClick={(e) => { e.stopPropagation(); onStartEdit(); }} 
                className="p-1 hover:bg-editorial-border rounded transition-colors text-editorial-muted hover:text-editorial-ink"
                title="Edit Title"
              >
                <Edit3 className="w-3 h-3" />
              </button>
            </div>
            {topic.sentiment && (
              <div className={cn("w-2 h-2 rounded-full shrink-0", sentimentColors[topic.sentiment])} />
            )}
          </div>
        </div>
      )}
    </div>
  );
}


