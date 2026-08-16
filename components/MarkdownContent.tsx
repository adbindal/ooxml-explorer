import React from 'react';
import ReactMarkdown from 'react-markdown';

interface MarkdownContentProps {
  content: string;
  className?: string;
}

/**
 * Renders AI-generated markdown (headings, bold, inline code, lists, fenced code
 * blocks, links) as actual formatted content instead of raw text. react-markdown
 * renders straight to React elements - no dangerouslySetInnerHTML, so there's no
 * hand-rolled escaping to get right even though model output routinely contains
 * literal `<w:tag>`-style XML.
 */
const MarkdownContent: React.FC<MarkdownContentProps> = ({ content, className = '' }) => {
  return (
    <div className={`space-y-2 [&_p]:leading-relaxed [&_ul]:list-disc [&_ul]:list-inside [&_ol]:list-decimal [&_ol]:list-inside [&_li]:leading-relaxed [&_h1]:text-sm [&_h1]:font-bold [&_h2]:text-sm [&_h2]:font-bold [&_h3]:text-xs [&_h3]:font-bold [&_strong]:text-blue-500 [&_strong]:font-bold [&_code]:bg-blue-500/10 [&_code]:px-1 [&_code]:py-0.5 [&_code]:rounded [&_code]:text-[10px] [&_code]:font-mono [&_code]:text-blue-500 [&_code]:border [&_code]:border-blue-500/20 [&_pre]:bg-black/20 [&_pre]:border [&_pre]:border-blue-500/10 [&_pre]:rounded-lg [&_pre]:p-3 [&_pre]:overflow-x-auto [&_pre_code]:bg-transparent [&_pre_code]:border-0 [&_pre_code]:p-0 [&_a]:text-blue-500 [&_a]:underline ${className}`}>
      <ReactMarkdown>{content}</ReactMarkdown>
    </div>
  );
};

export default MarkdownContent;
