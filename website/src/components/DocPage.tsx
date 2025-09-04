import React from 'react';

type TocItem = { id: string; label: string };

export default function DocPage({
  title,
  toc,
  children,
}: {
  title: string;
  toc?: TocItem[];
  children: React.ReactNode;
}): JSX.Element {
  return (
    <article className="doc">
      <h1>{title}</h1>
      {toc && toc.length > 0 && (
        <nav className="toc" aria-label="Table of contents">
          <div className="toc-title">On this page</div>
          <ul>
            {toc.map(item => (
              <li key={item.id}>
                <a href={`#${item.id}`}>{item.label}</a>
              </li>
            ))}
          </ul>
        </nav>
      )}
      {children}
    </article>
  );
}


