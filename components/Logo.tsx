import React from 'react';

interface LogoProps {
  className?: string;
  size?: number;
  theme?: 'dark' | 'light';
}

const Logo: React.FC<LogoProps> = ({ className = "", size = 40, theme = 'light' }) => {
  // Brand Colors
  const primary = '#4A89DC'; // Light Blue
  const text = theme === 'dark' ? '#E2E8F0' : '#1F3F70'; 
  
  return (
    <svg 
      width={size} 
      height={size} 
      viewBox="0 0 64 64" 
      fill="none" 
      xmlns="http://www.w3.org/2000/svg"
      className={className}
      role="img"
      aria-label="OOXML Explorer Logo"
    >
      <title>OOXML Explorer Logo</title>
      
      {/* Main Document Shape */}
      <path 
        d="M14 6H40L52 18V58H14V6Z" 
        fill={theme === 'dark' ? 'rgba(74, 137, 220, 0.15)' : '#F0F9FF'} 
        stroke={primary} 
        strokeWidth="3" 
        strokeLinejoin="round"
      />
      
      {/* Folded Corner */}
      <path 
        d="M40 6V18H52" 
        fill={theme === 'dark' ? 'rgba(74, 137, 220, 0.3)' : '#DBEAFE'}
        stroke={primary} 
        strokeWidth="3" 
        strokeLinejoin="round" 
        strokeLinecap="round"
      />
      
      {/* Code Brackets Symbol < > */}
      <g transform="translate(33, 38) scale(0.8)">
        <path 
          d="M-8 -6L-14 0L-8 6" 
          stroke={text} 
          strokeWidth="4" 
          strokeLinecap="round" 
          strokeLinejoin="round"
        />
        <path 
          d="M8 -6L14 0L8 6" 
          stroke={text} 
          strokeWidth="4" 
          strokeLinecap="round" 
          strokeLinejoin="round"
        />
        <path 
          d="M-3 8L3 -8" 
          stroke={primary} 
          strokeWidth="3" 
          strokeLinecap="round" 
          strokeLinejoin="round"
        />
      </g>
    </svg>
  );
};

export default Logo;