export interface SourceFile {
  name: string;
  size: number;
  type: string;
  lastModified: number;
  topicCount: number;
}

export interface BookTopic {
  id: string;
  title: string;
  summary: string;
  content: string;
  sourceFile: string;
  tags: string[];
  sentiment?: 'analytical' | 'creative' | 'technical' | 'narrative';
  wordCount: number;
  processSteps?: { title: string; description: string }[];
}

export interface Book {
  title: string;
  topics: BookTopic[];
}

export interface FileData {
  name: string;
  content: string;
  type: string;
}
