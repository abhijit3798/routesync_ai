import React, { useState, useMemo } from 'react';
import { Search, ChevronLeft, ChevronRight, Download, Edit2, Check, X, Plus } from 'lucide-react';
import { logger } from '../utils/logger';

interface DataTableProps {
  data: Record<string, unknown>[];
  sheetName: string;
  onDataChange: (updatedData: Record<string, unknown>[]) => void;
  onExport: () => void;
}

export const DataTable: React.FC<DataTableProps> = ({
  data,
  sheetName,
  onDataChange,
  onExport,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [sortConfig, setSortConfig] = useState<{ key: string; direction: 'asc' | 'desc' } | null>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  
  // Cell editing state
  const [editingCell, setEditingCell] = useState<{ rowIndex: number; columnKey: string } | null>(null);
  const [editValue, setEditValue] = useState<string>('');

  // Extract columns dynamically from data keys
  const columns = useMemo(() => {
    if (!data || data.length === 0) return [];
    
    // Collect all unique keys from all rows (handles sparse rows)
    const allKeys = new Set<string>();
    data.forEach(row => {
      Object.keys(row).forEach(key => allKeys.add(key));
    });
    return Array.from(allKeys);
  }, [data]);

  // Handle Sort
  const requestSort = (key: string) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
    logger.info(`Sorting columns by ${key} [${direction}]`);
  };

  // Filter & Sort data
  const processedData = useMemo(() => {
    let result = [...data];

    // Filter
    if (searchTerm.trim() !== '') {
      const lowerSearch = searchTerm.toLowerCase();
      result = result.filter((row) =>
        Object.values(row).some((val) =>
          String(val).toLowerCase().includes(lowerSearch)
        )
      );
    }

    // Sort
    if (sortConfig) {
      const { key, direction } = sortConfig;
      result.sort((a, b) => {
        const valA = a[key] ?? '';
        const valB = b[key] ?? '';

        if (valA < valB) return direction === 'asc' ? -1 : 1;
        if (valA > valB) return direction === 'asc' ? 1 : -1;
        return 0;
      });
    }

    return result;
  }, [data, searchTerm, sortConfig]);

  // Pagination calculations
  const totalRows = processedData.length;
  const totalPages = Math.ceil(totalRows / rowsPerPage);
  
  const currentRows = useMemo(() => {
    const startIdx = (currentPage - 1) * rowsPerPage;
    return processedData.slice(startIdx, startIdx + rowsPerPage);
  }, [processedData, currentPage, rowsPerPage]);

  // Adjust page if current rows decrease
  React.useEffect(() => {
    if (currentPage > 1 && currentRows.length === 0) {
      setCurrentPage(Math.max(1, totalPages));
    }
  }, [currentRows, currentPage, totalPages]);

  // Cell edit actions
  const startEditing = (rowIndex: number, columnKey: string, currentValue: unknown) => {
    setEditingCell({ rowIndex, columnKey });
    setEditValue(String(currentValue ?? ''));
  };

  const saveEdit = (originalIndex: number, columnKey: string) => {
    if (!editingCell) return;
    
    const updated = [...data];
    updated[originalIndex] = {
      ...updated[originalIndex],
      [columnKey]: editValue,
    };
    
    onDataChange(updated);
    setEditingCell(null);
    logger.info(`Cell at row ${originalIndex}, col ${columnKey} updated to: ${editValue}`);
  };

  const cancelEdit = () => {
    setEditingCell(null);
  };

  // Add a new empty row
  const handleAddRow = () => {
    const newRow: Record<string, unknown> = {};
    columns.forEach(col => {
      newRow[col] = '';
    });
    const updated = [newRow, ...data];
    onDataChange(updated);
    setCurrentPage(1); // Jump to first page
    logger.info('Added new empty row to top of the table');
  };

  return (
    <div className="glass-panel animate-fade-in" style={tableCardStyle}>
      {/* Table Header / Action Area */}
      <div style={actionHeaderStyle}>
        <div style={titleAreaStyle}>
          <h2 style={sheetTitleStyle}>{sheetName}</h2>
          <span style={sheetBadgeStyle}>{data.length} records</span>
        </div>

        <div style={controlsAreaStyle}>
          {/* Search bar */}
          <div style={searchWrapperStyle}>
            <Search size={16} color="var(--text-muted)" style={searchIconStyle} />
            <input
              id="table-search-input"
              type="text"
              placeholder="Search in dataset..."
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setCurrentPage(1);
              }}
              style={searchInputStyle}
            />
          </div>

          {/* Action buttons */}
          <button id="table-add-row-btn" onClick={handleAddRow} style={secondaryButtonStyle}>
            <Plus size={16} style={{ marginRight: '6px' }} />
            Add Row
          </button>

          <button id="table-export-btn" onClick={onExport} style={primaryButtonStyle}>
            <Download size={16} style={{ marginRight: '6px' }} />
            Export Excel
          </button>
        </div>
      </div>

      {/* Table Container */}
      <div style={tableContainerStyle}>
        {currentRows.length === 0 ? (
          <div style={noResultsStyle}>No matches found. Try relaxing your filters.</div>
        ) : (
          <table style={tableStyle}>
            <thead>
              <tr style={tableHeaderRowStyle}>
                <th style={{ ...tableHeaderCellStyle, width: '60px' }}>#</th>
                {columns.map((col) => (
                  <th
                    key={col}
                    onClick={() => requestSort(col)}
                    style={tableHeaderCellStyle}
                  >
                    <div style={headerCellInnerStyle}>
                      <span>{col}</span>
                      {sortConfig?.key === col ? (
                        <span style={sortIndicatorStyle}>
                          {sortConfig.direction === 'asc' ? ' ▲' : ' ▼'}
                        </span>
                      ) : (
                        <span style={{ ...sortIndicatorStyle, opacity: 0.2 }}> ↕</span>
                      )}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {currentRows.map((row, relativeIdx) => {
                // Calculate absolute index in primary array
                const absoluteIdx = (currentPage - 1) * rowsPerPage + relativeIdx;
                
                return (
                  <tr key={absoluteIdx} style={tableBodyRowStyle}>
                    <td style={tableBodyCellStyle_index}>{absoluteIdx + 1}</td>
                    {columns.map((col) => {
                      const isEditing = editingCell?.rowIndex === absoluteIdx && editingCell?.columnKey === col;
                      const value = row[col];

                      return (
                        <td
                          key={col}
                          style={tableBodyCellStyle}
                          onDoubleClick={() => !isEditing && startEditing(absoluteIdx, col, value)}
                        >
                          {isEditing ? (
                            <div style={editWrapperStyle}>
                              <input
                                type="text"
                                value={editValue}
                                onChange={(e) => setEditValue(e.target.value)}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') saveEdit(absoluteIdx, col);
                                  if (e.key === 'Escape') cancelEdit();
                                }}
                                style={editInputStyle}
                                autoFocus
                              />
                              <div style={editButtonsStyle}>
                                <button
                                  onClick={() => saveEdit(absoluteIdx, col)}
                                  style={actionIconButton_check}
                                >
                                  <Check size={14} />
                                </button>
                                <button onClick={cancelEdit} style={actionIconButton_cancel}>
                                  <X size={14} />
                                </button>
                              </div>
                            </div>
                          ) : (
                            <div style={cellValueContainerStyle}>
                              <span style={cellTextStyle}>{String(value ?? '')}</span>
                              <button
                                onClick={() => startEditing(absoluteIdx, col, value)}
                                className="cell-edit-btn"
                                style={cellEditButtonStyle}
                              >
                                <Edit2 size={12} />
                              </button>
                            </div>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      {/* Pagination Footer */}
      <div style={footerStyle}>
        <div style={footerLeftStyle}>
          <span style={footerTextStyle}>Rows per page:</span>
          <select
            id="table-page-size-select"
            value={rowsPerPage}
            onChange={(e) => {
              setRowsPerPage(Number(e.target.value));
              setCurrentPage(1);
            }}
            style={selectStyle}
          >
            {[5, 10, 25, 50, 100].map((size) => (
              <option key={size} value={size}>
                {size}
              </option>
            ))}
          </select>
          <span style={footerRangeTextStyle}>
            Showing {Math.min(totalRows, (currentPage - 1) * rowsPerPage + 1)} -{' '}
            {Math.min(totalRows, currentPage * rowsPerPage)} of {totalRows}
          </span>
        </div>

        <div style={paginationButtonsStyle}>
          <button
            id="table-prev-page-btn"
            onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
            disabled={currentPage === 1}
            style={currentPage === 1 ? disabledPageButtonStyle : pageButtonStyle}
          >
            <ChevronLeft size={16} />
          </button>
          <span style={pageNumberTextStyle}>
            Page {currentPage} of {totalPages || 1}
          </span>
          <button
            id="table-next-page-btn"
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            disabled={currentPage === totalPages || totalPages === 0}
            style={currentPage === totalPages || totalPages === 0 ? disabledPageButtonStyle : pageButtonStyle}
          >
            <ChevronRight size={16} />
          </button>
        </div>
      </div>
    </div>
  );
};

// Styles
const tableCardStyle: React.CSSProperties = {
  width: '100%',
  display: 'flex',
  flexDirection: 'column',
  padding: '0',
  overflow: 'hidden',
  background: 'var(--bg-surface)',
  border: '1px solid var(--border-color)',
};

const actionHeaderStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  flexWrap: 'wrap',
  gap: '16px',
  padding: '20px 24px',
  borderBottom: '1px solid var(--border-color)',
};

const titleAreaStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: '12px',
};

const sheetTitleStyle: React.CSSProperties = {
  fontSize: '1.25rem',
  fontWeight: 600,
  color: 'var(--text-primary)',
};

const sheetBadgeStyle: React.CSSProperties = {
  fontSize: '0.75rem',
  fontWeight: 600,
  backgroundColor: 'var(--primary-glow)',
  color: 'var(--primary)',
  padding: '4px 10px',
  borderRadius: 'var(--radius-full)',
};

const controlsAreaStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: '12px',
  flexWrap: 'wrap',
};

const searchWrapperStyle: React.CSSProperties = {
  position: 'relative',
  display: 'flex',
  alignItems: 'center',
  maxWidth: '240px',
};

const searchIconStyle: React.CSSProperties = {
  position: 'absolute',
  left: '12px',
  pointerEvents: 'none',
};

const searchInputStyle: React.CSSProperties = {
  paddingLeft: '36px',
  paddingTop: '8px',
  paddingBottom: '8px',
  fontSize: '0.85rem',
  width: '100%',
  borderRadius: '6px',
};

const primaryButtonStyle: React.CSSProperties = {
  backgroundColor: 'var(--primary)',
  color: 'white',
  border: 'none',
  padding: '8px 16px',
  borderRadius: '6px',
  fontSize: '0.875rem',
  fontWeight: 600,
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  boxShadow: 'var(--shadow-sm)',
  transition: 'background-color 0.2s',
};

const secondaryButtonStyle: React.CSSProperties = {
  backgroundColor: 'transparent',
  color: 'var(--text-primary)',
  border: '1px solid var(--border-color)',
  padding: '8px 16px',
  borderRadius: '6px',
  fontSize: '0.875rem',
  fontWeight: 600,
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  transition: 'background-color 0.2s, border-color 0.2s',
};

const tableContainerStyle: React.CSSProperties = {
  width: '100%',
  maxHeight: '500px',
  overflow: 'auto',
  position: 'relative',
};

const tableStyle: React.CSSProperties = {
  width: '100%',
  borderCollapse: 'collapse',
  textAlign: 'left',
};

const tableHeaderRowStyle: React.CSSProperties = {
  position: 'sticky',
  top: 0,
  zIndex: 10,
  backgroundColor: 'var(--bg-surface-elevated)',
  boxShadow: '0 1px 0 var(--border-color)',
};

const tableHeaderCellStyle: React.CSSProperties = {
  padding: '14px 16px',
  fontSize: '0.8rem',
  fontWeight: 600,
  color: 'var(--text-secondary)',
  textTransform: 'uppercase',
  letterSpacing: '0.05em',
  borderBottom: '1px solid var(--border-color)',
  cursor: 'pointer',
  userSelect: 'none',
  whiteSpace: 'nowrap',
};

const headerCellInnerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
};

const sortIndicatorStyle: React.CSSProperties = {
  marginLeft: '4px',
  fontSize: '0.85rem',
};

const tableBodyRowStyle: React.CSSProperties = {
  borderBottom: '1px solid var(--border-color)',
  transition: 'background-color var(--transition-fast)',
};

// Target hover selector via custom inline styling is limited, but we handle it smoothly
const tableBodyCellStyle: React.CSSProperties = {
  padding: '12px 16px',
  fontSize: '0.875rem',
  color: 'var(--text-primary)',
  maxWidth: '220px',
  whiteSpace: 'nowrap',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  position: 'relative',
};

const tableBodyCellStyle_index: React.CSSProperties = {
  ...tableBodyCellStyle,
  color: 'var(--text-muted)',
  fontWeight: 500,
};

const noResultsStyle: React.CSSProperties = {
  padding: '48px',
  textAlign: 'center',
  color: 'var(--text-secondary)',
  fontSize: '0.95rem',
};

const cellValueContainerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  width: '100%',
};

const cellTextStyle: React.CSSProperties = {
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  marginRight: '8px',
};

const cellEditButtonStyle: React.CSSProperties = {
  background: 'none',
  border: 'none',
  color: 'var(--text-muted)',
  cursor: 'pointer',
  opacity: 0,
  transition: 'opacity 0.2s, color 0.2s',
  padding: '2px',
  borderRadius: '3px',
};

const editWrapperStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: '8px',
  width: '100%',
};

const editInputStyle: React.CSSProperties = {
  flexGrow: 1,
  padding: '4px 8px',
  fontSize: '0.85rem',
  borderRadius: '4px',
  width: '100%',
};

const editButtonsStyle: React.CSSProperties = {
  display: 'flex',
  gap: '4px',
};

const actionIconButton_check: React.CSSProperties = {
  background: 'var(--success-glow)',
  border: '1px solid var(--success)',
  borderRadius: '4px',
  color: 'var(--success)',
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  padding: '4px',
};

const actionIconButton_cancel: React.CSSProperties = {
  background: 'var(--error-glow)',
  border: '1px solid var(--error)',
  borderRadius: '4px',
  color: 'var(--error)',
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  padding: '4px',
};

const footerStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  padding: '16px 24px',
  borderTop: '1px solid var(--border-color)',
  flexWrap: 'wrap',
  gap: '16px',
};

const footerLeftStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: '12px',
};

const footerTextStyle: React.CSSProperties = {
  fontSize: '0.85rem',
  color: 'var(--text-secondary)',
};

const selectStyle: React.CSSProperties = {
  padding: '4px 8px',
  fontSize: '0.85rem',
  borderRadius: '4px',
};

const footerRangeTextStyle: React.CSSProperties = {
  fontSize: '0.85rem',
  color: 'var(--text-muted)',
  marginLeft: '12px',
};

const paginationButtonsStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: '8px',
};

const pageButtonStyle: React.CSSProperties = {
  background: 'var(--bg-surface-elevated)',
  border: '1px solid var(--border-color)',
  borderRadius: '6px',
  color: 'var(--text-primary)',
  cursor: 'pointer',
  padding: '6px',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
};

const disabledPageButtonStyle: React.CSSProperties = {
  ...pageButtonStyle,
  opacity: 0.4,
  cursor: 'not-allowed',
};

const pageNumberTextStyle: React.CSSProperties = {
  fontSize: '0.85rem',
  color: 'var(--text-secondary)',
  margin: '0 8px',
};

// CSS injected to display edit buttons on row hover
if (typeof document !== 'undefined') {
  const style = document.createElement('style');
  style.innerHTML = `
    tbody tr:hover {
      background-color: var(--primary-glow);
    }
    tbody tr:hover .cell-edit-btn {
      opacity: 1 !important;
    }
    .cell-edit-btn:hover {
      color: var(--primary) !important;
      background-color: var(--border-color);
    }
  `;
  document.head.appendChild(style);
}
