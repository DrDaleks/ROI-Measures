package plugins.adufour.roi;

import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.event.TableModelListener;
import javax.swing.table.TableModel;

import jxl.Cell;
import jxl.Sheet;
import jxl.write.WritableSheet;

/**
 * Excel table view
 * 
 * @author Fabrice de Chaumont, Alexandre Dufour
 * 
 */
public class ExcelTable extends JScrollPane
{
	private static final long	serialVersionUID	= 1L;
	
	private JTable				table;
	
	public ExcelTable()
	{
		
	}
	
	public ExcelTable(WritableSheet page)
	{
		updateSheet(page);
		this.setViewportView(table);
		this.setAutoscrolls(true);
	}
	
	public synchronized void updateSheet(WritableSheet page)
	{
		table = new JTable();
		this.setViewportView(table);
		if (page != null)
		{
			table.setModel(new SheetTableModel(page));
		}
	}
	
	private class SheetTableModel implements TableModel
	{
		private Sheet	sheet	= null;
		
		public SheetTableModel(Sheet sheet)
		{
			this.sheet = sheet;
		}
		
		public int getRowCount()
		{
			return sheet.getRows();
		}
		
		public int getColumnCount()
		{
			
			return sheet.getColumns();
		}
		
		/**
		 * Copied from javax.swing.table.AbstractTableModel, to name columns using spreadsheet
		 * conventions: A, B, C, . Z, AA, AB, etc.
		 */
		public String getColumnName(int column)
		{
			String result = "";
			for (; column >= 0; column = column / 26 - 1)
			{
				result = (char) ((char) (column % 26) + 'A') + result;
			}
			return result;
		}
		
		public Class<?> getColumnClass(int columnIndex)
		{
			return String.class;
		}
		
		public boolean isCellEditable(int rowIndex, int columnIndex)
		{
			return false;
		}
		
		public Object getValueAt(int rowIndex, int columnIndex)
		{
			
			try
			{
				Cell cell = sheet.getCell(columnIndex, rowIndex);
				return cell.getContents();
			}
			catch (Exception e)
			{
				//
			}
			return null;
		}
		
		public void setValueAt(Object aValue, int rowIndex, int columnIndex)
		{
		}
		
		public void addTableModelListener(TableModelListener l)
		{
		}
		
		public void removeTableModelListener(TableModelListener l)
		{
		}
		
	}
}
