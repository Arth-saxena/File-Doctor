'use client';
import { Home, ChevronRight } from 'lucide-react';

interface BreadcrumbItem {
  label: string;
  href?: string;
}

interface BreadcrumbHeaderProps {
  title: string;
  subtitle?: string;
  items: BreadcrumbItem[];
  actions?: React.ReactNode;
}

export default function BreadcrumbHeader({ title, subtitle, items, actions }: BreadcrumbHeaderProps) {
  return (
    <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm mb-6 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
      {/* Left: Trail & Page Title */}
      <div className="space-y-1">
        <nav className="flex items-center space-x-1.5 text-xs text-slate-500 font-medium">
          <a href="#" className="hover:text-blue-600 transition flex items-center space-x-1">
            <Home className="w-3.5 h-3.5 text-slate-400" />
            <span>App</span>
          </a>
          
          {items.map((item, index) => (
            <div key={index} className="flex items-center space-x-1.5">
              <ChevronRight className="w-3 h-3 text-slate-400 flex-shrink-0" />
              {item.href ? (
                <a href={item.href} className="hover:text-blue-600 transition">
                  {item.label}
                </a>
              ) : (
                <span className="text-slate-800 font-semibold">{item.label}</span>
              )}
            </div>
          ))}
        </nav>

        <div className="flex items-center space-x-3">
          <h1 className="text-2xl font-bold text-slate-900 tracking-tight">{title}</h1>
        </div>
        
        {subtitle && (
          <p className="text-xs text-slate-500">{subtitle}</p>
        )}
      </div>

      {/* Right: APEX Action Buttons Slot */}
      {actions && (
        <div className="flex items-center space-x-3 flex-shrink-0">
          {actions}
        </div>
      )}
    </div>
  );
}