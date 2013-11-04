package plugins.adufour.roi;

import icy.gui.main.GlobalSequenceListener;
import icy.main.Icy;
import icy.roi.BooleanMask2D;
import icy.roi.ROI;
import icy.roi.ROI2D;
import icy.roi.ROI3D;
import icy.sequence.Sequence;
import icy.sequence.SequenceEvent;
import icy.sequence.SequenceEvent.SequenceEventSourceType;
import icy.sequence.SequenceListener;
import icy.system.thread.ThreadUtil;
import icy.type.collection.array.Array1DUtil;
import icy.type.rectangle.Rectangle3D;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.TreeMap;
import java.util.UUID;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import javax.vecmath.Point3d;

import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.RGB;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import plugins.adufour.ezplug.EzButton;
import plugins.adufour.ezplug.EzException;
import plugins.adufour.ezplug.EzPlug;
import plugins.adufour.ezplug.EzVar;
import plugins.adufour.ezplug.EzVarBoolean;
import plugins.adufour.ezplug.EzVarInteger;
import plugins.adufour.ezplug.EzVarListener;
import plugins.adufour.ezplug.EzVarSequence;

public class ROIMeasures extends EzPlug implements GlobalSequenceListener, SequenceListener
{
    private static File                xlsFile;
    
    private static WritableWorkbook    workbook;
    
    // Create a static workbook instance to make sure a single file is opened
    static
    {
        try
        {
            xlsFile = File.createTempFile("icy_ROI_Measures_" + UUID.randomUUID().toString(), "xls");
            workbook = Workbook.createWorkbook(xlsFile);
            
        }
        catch (IOException e)
        {
            e.printStackTrace();
            throw new EzException("ROI Meaasures: unable to initialize\n(see output console for details)", true);
        }
    }
    
    private final EzVarSequence        currentSeq = new EzVarSequence("Current sequence");
    
    private final EzVarBoolean         restrictZ  = new EzVarBoolean("Z selection", false);
    private final EzVarInteger         selectedZ  = new EzVarInteger("Selected Z", 0, 0, 0, 1);
    
    private final EzVarBoolean         liveUpdate = new EzVarBoolean("Auto-update", true);
    
    private final ExcelTable           table      = new ExcelTable();
    
    private final HashMap<ROI, Color>  ROIColors  = new HashMap<ROI, Color>();
    private final HashMap<ROI, Colour> ROIColours = new HashMap<ROI, Colour>();
    
    @Override
    protected void initialize()
    {
        getUI().setParametersIOVisible(false);
        getUI().setRunButtonText("Update");
        getUI().setActionPanelVisible(!liveUpdate.getValue());
        
        addEzComponent(currentSeq);
        
        addEzComponent(restrictZ);
        
        restrictZ.addVisibilityTriggerTo(selectedZ, true);
        
        restrictZ.addVarChangeListener(new EzVarListener<Boolean>()
        {
            @Override
            public void variableChanged(EzVar<Boolean> source, Boolean newValue)
            {
                updateMaxZ();
            }
        });
        
        addEzComponent(selectedZ);
        
        selectedZ.addVarChangeListener(new EzVarListener<Integer>()
        {
            @Override
            public void variableChanged(EzVar<Integer> source, Integer newValue)
            {
                execute();
            }
        });
        
        addEzComponent(liveUpdate);
        
        table.setPreferredSize(new Dimension(500, 100));
        
        addComponent(table);
        
        try
        {
            xlsFile = File.createTempFile("icy_ROI_Measures_" + UUID.randomUUID().toString(), "xls");
            workbook = Workbook.createWorkbook(xlsFile);
            
        }
        catch (IOException e)
        {
            e.printStackTrace();
            throw new EzException("ROI Meaasures: unable to initialize\n(see output console for details)", true);
        }
        
        EzButton buttonExport = new EzButton("Export to .CSV file...", new ActionListener()
        {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                JFileChooser jfc = new JFileChooser();
                jfc.setFileFilter(new FileFilter()
                {
                    @Override
                    public boolean accept(File f)
                    {
                        return f.isDirectory() || f.getAbsolutePath().toLowerCase().endsWith(".csv");
                    }
                    
                    @Override
                    public String getDescription()
                    {
                        return ".CSV (Comma Separated Values)";
                    }
                });
                
                if (jfc.showSaveDialog(null) == JFileChooser.APPROVE_OPTION)
                {
                    try
                    {
                        ExportCSV.export(workbook, jfc.getSelectedFile(), false);
                    }
                    catch (IOException e1)
                    {
                        e1.printStackTrace();
                    }
                }
            }
        });
        
        addEzComponent(buttonExport);
        
        // Listeners
        
        // add a listener to currently opened sequences
        for (final Sequence sequence : Icy.getMainInterface().getSequences())
            sequence.addListener(this);
        
        // add a listener to detect newly opened sequences
        Icy.getMainInterface().addGlobalSequenceListener(this);
        
        currentSeq.addVarChangeListener(new EzVarListener<Sequence>()
        {
            @Override
            public void variableChanged(EzVar<Sequence> source, Sequence sequence)
            {
                if (sequence == null || sequence.getSizeZ() == 1)
                {
                    restrictZ.setVisible(false);
                    restrictZ.setValue(false);
                    selectedZ.setValue(0);
                }
                else
                {
                    restrictZ.setVisible(true);
                }
                
                table.updateSheet(sequence == null ? null : getOrCreateSheet(sequence));
                if (sequence != null && sequence.getFirstViewer() != null) update(sequence);
                getUI().repack(false);
            }
        });
        
        liveUpdate.addVarChangeListener(new EzVarListener<Boolean>()
        {
            @Override
            public void variableChanged(EzVar<Boolean> source, Boolean newValue)
            {
                getUI().setActionPanelVisible(!newValue);
            }
        });
        
        try
        {
            currentSeq.setValue(getActiveSequence());
            update(currentSeq.getValue());
            table.repaint();
        }
        catch (EzException e)
        {
        }
    }
    
    private void updateMaxZ()
    {
        if (currentSeq.getValue() != null)
        {
            int maxZ = currentSeq.getValue().getSizeZ() - 1;
            
            selectedZ.setMaxValue(maxZ);
            
            if (selectedZ.getValue() >= maxZ)
            {
                selectedZ.setValue(maxZ);
            }
            else
            {
                if (liveUpdate.getValue()) execute();
            }
        }
    }
    
    /**
     * Creates a new sheet for the given sequence, or returns the current existing sheet (if any)
     * 
     * @param sequence
     * @return
     */
    private static synchronized WritableSheet getOrCreateSheet(Sequence sequence)
    {
        String sheetName = getSheetName(sequence);
        
        WritableSheet sheet = workbook.getSheet(sheetName);
        
        if (sheet == null)
        {
            sheet = workbook.createSheet(sheetName, workbook.getNumberOfSheets() + 1);
            
            try
            {
                sheet.addCell(new Label(0, 0, "Name"));
                sheet.addCell(new Label(1, 0, "Color"));
                sheet.addCell(new Label(2, 0, "Min. X"));
                sheet.addCell(new Label(3, 0, "Min. Y"));
                sheet.addCell(new Label(4, 0, "Width"));
                sheet.addCell(new Label(5, 0, "Height"));
                sheet.addCell(new Label(6, 0, "Perimeter"));
                sheet.addCell(new Label(7, 0, "Area"));
                
                for (int c = 0; c < sequence.getSizeC(); c++)
                {
                    sheet.addCell(new Label(7 + 4 * c + 1, 0, "Min. (" + sequence.getChannelName(c) + ")"));
                    sheet.addCell(new Label(7 + 4 * c + 2, 0, "Avg. (" + sequence.getChannelName(c) + ")"));
                    sheet.addCell(new Label(7 + 4 * c + 3, 0, "Max. (" + sequence.getChannelName(c) + ")"));
                    sheet.addCell(new Label(7 + 4 * c + 4, 0, "Sum. (" + sequence.getChannelName(c) + ")"));
                }
                
            }
            catch (RowsExceededException e)
            {
                e.printStackTrace();
            }
            catch (WriteException e)
            {
                e.printStackTrace();
            }
        }
        
        return sheet;
    }
    
    private static String getSheetName(Sequence s)
    {
        return "Sequence " + s.hashCode();
    }
    
    @Override
    protected void execute()
    {
        update(currentSeq.getValue());
    }
    
    private void update(Sequence sequence)
    {
        if (sequence == null) return;
        
        for (ROI roi : sequence.getROIs())
            update(sequence, roi);
    }
    
    private void update(Sequence sequence, final ROI roi)
    {
        try
        {
            int row = sequence.getROIs().indexOf(roi) + 1;
            
            WritableSheet sheet = getOrCreateSheet(sequence);
            
            // set the name
            sheet.addCell(new Label(0, row, roi.getName()));
            
            // set the color (if it has changed)
            if (roi.getColor() != ROIColors.get(roi))
            {
                
                // JXL is VERY nasty, won't allow custom (AWT) colors !!!
                
                // walk-around: look for the closest color match
                Color col = roi.getColor();
                Point3d cval = new Point3d(col.getRed(), col.getGreen(), col.getBlue());
                
                Colour colour = Colour.BLACK;
                double minDistance = Double.MAX_VALUE;
                for (Colour c : Colour.getAllColours())
                {
                    RGB rgb = c.getDefaultRGB();
                    Point3d cval_jxl = new Point3d(rgb.getRed(), rgb.getGreen(), rgb.getBlue());
                    
                    // compute (squared) distance (to go faster)
                    // I mean, this is lame anyway, isn't it !!
                    double dist = cval.distanceSquared(cval_jxl);
                    if (dist < minDistance)
                    {
                        minDistance = dist;
                        colour = c;
                    }
                }
                
                // do you realize what we did JUST to get the color ?!!
                ROIColors.put(roi, roi.getColor());
                ROIColours.put(roi, colour);
            }
            
            sheet.addCell(new Label(1, row, ROIColours.get(roi).getDescription(), new WritableCellFormat()
            {
                @Override
                public Colour getBackgroundColour()
                {
                    return ROIColours.get(roi);
                }
            }));
            
            double[] min = new double[sequence.getSizeC()];
            double[] max = new double[sequence.getSizeC()];
            double[] sum = new double[sequence.getSizeC()];
            double[] cpt = new double[sequence.getSizeC()];
            
            if (roi instanceof ROI2D)
            {
                ROI2D r2 = (ROI2D) roi;
                
                // set x,y,w,h
                Rectangle bounds = r2.getBounds();
                sheet.addCell(new Number(2, row, bounds.x));
                sheet.addCell(new Number(3, row, bounds.y));
                sheet.addCell(new Number(4, row, bounds.width));
                sheet.addCell(new Number(5, row, bounds.height));
                sheet.addCell(new Number(6, row, r2.getPerimeter()));
                sheet.addCell(new Number(7, row, r2.getNumberOfPoints()));
                
                boolean[] mask = r2.getBooleanMask(true).mask;
                Object[][] z_c_xy = (Object[][]) sequence.getDataXYCZ(sequence.getFirstViewer().getPositionT());
                boolean signed = sequence.getDataType_().isSigned();
                int width = sequence.getSizeX();
                int height = sequence.getSizeY();
                
                int ioff = bounds.x + bounds.y * width;
                int moff = 0;
                
                int minZ = restrictZ.getValue() ? selectedZ.getValue() : 0;
                int maxZ = restrictZ.getValue() ? minZ : sequence.getSizeZ(sequence.getFirstViewer().getPositionT()) - 1;
                
                for (int iy = bounds.y, my = 0; my < bounds.height; my++, iy++, ioff += sequence.getSizeX() - bounds.width)
                    for (int ix = bounds.x, mx = 0; mx < bounds.width; mx++, ix++, ioff++, moff++)
                    {
                        if (iy >= 0 && ix >= 0 && iy < height && ix < width && mask[moff])
                        {
                            for (int z = minZ; z <= maxZ; z++)
                                for (int c = 0; c < sum.length; c++)
                                {
                                    cpt[c]++;
                                    double val = Array1DUtil.getValue(z_c_xy[z][c], ioff, signed);
                                    sum[c] += val;
                                    if (val > max[c]) max[c] = val;
                                    if (val < min[c]) min[c] = val;
                                }
                        }
                    }
            }
            else if (roi instanceof ROI3D)
            {
                ROI3D r3 = (ROI3D) roi;
                
                // set x,y,w,h
                Rectangle3D.Integer bounds3 = r3.getBounds();
                sheet.addCell(new Number(2, row, bounds3.x));
                sheet.addCell(new Number(3, row, bounds3.y));
                sheet.addCell(new Number(4, row, bounds3.getSizeX()));
                sheet.addCell(new Number(5, row, bounds3.getSizeY()));
                sheet.addCell(new Number(6, row, r3.getNumberOfContourPoints()));
                sheet.addCell(new Number(7, row, r3.getNumberOfPoints()));
                
                TreeMap<Integer, BooleanMask2D> masks = r3.getBooleanMask(true).mask;
                Object[][] z_c_xy = (Object[][]) sequence.getDataXYCZ(sequence.getFirstViewer().getPositionT());
                boolean signed = sequence.getDataType_().isSigned();
                int width = sequence.getSizeX();
                int height = sequence.getSizeY();
                
                for (Integer z : masks.keySet())
                {
                    boolean[] mask = masks.get(z).mask;
                    Rectangle bounds = masks.get(z).bounds;
                    int ioff = bounds.x + bounds.y * width;
                    int moff = 0;
                    
                    for (int iy = bounds.y, my = 0; my < bounds.height; my++, iy++, ioff += sequence.getSizeX() - bounds.width)
                        for (int ix = bounds.x, mx = 0; mx < bounds.width; mx++, ix++, ioff++, moff++)
                        {
                            if (iy >= 0 && ix >= 0 && iy < height && ix < width && mask[moff])
                            {
                                for (int c = 0; c < sum.length; c++)
                                {
                                    cpt[c]++;
                                    double val = Array1DUtil.getValue(z_c_xy[z][c], ioff, signed);
                                    sum[c] += val;
                                    if (val > max[c]) max[c] = val;
                                    if (val < min[c]) min[c] = val;
                                }
                            }
                        }
                }
            }
            
            for (int c = 0; c < sum.length; c++)
            {
                sheet.addCell(new Number(7 + 4 * c + 1, row, min[c]));
                sheet.addCell(new Number(7 + 4 * c + 2, row, sum[c] / cpt[c]));
                sheet.addCell(new Number(7 + 4 * c + 3, row, max[c]));
                sheet.addCell(new Number(7 + 4 * c + 4, row, sum[c]));
            }
            
        }
        catch (Exception e)
        {
            e.printStackTrace();
            // We want a bug report to investigate the issue
            throw new RuntimeException(e);
        }
    }
    
    @Override
    public void clean()
    {
        ROIColors.clear();
        
        // remove listeners
        for (final Sequence sequence : Icy.getMainInterface().getSequences())
            sequence.removeListener(this);
        
        Icy.getMainInterface().removeGlobalSequenceListener(this);
        
        // clean static stuff (if this is the last instance)
        if (getNbInstances() == 1)
        {
            try
            {
                workbook.close();
                xlsFile.deleteOnExit();
            }
            catch (WriteException e)
            {
                e.printStackTrace();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }
    
    // MAIN LISTENER //
    
    @Override
    public void sequenceOpened(Sequence sequence)
    {
        sequence.addListener(this);
    }
    
    @Override
    public void sequenceClosed(Sequence sequence)
    {
        sequence.removeListener(this);
    }
    
    // SEQUENCE LISTENER //
    
    @Override
    public void sequenceChanged(SequenceEvent sequenceEvent)
    {
        Sequence sequence = sequenceEvent.getSequence();
        
        Sequence selected = currentSeq.getValue();
        if (selected != sequence) return;
        
        if (sequenceEvent.getSourceType() == SequenceEventSourceType.SEQUENCE_DATA)
        {
            ThreadUtil.invokeLater(new Runnable()
            {
                
                @Override
                public void run()
                {
                    updateMaxZ();
                }
            });
        }
        
        if (!liveUpdate.getValue()) return;
        
        if (sequenceEvent.getSourceType() == SequenceEventSourceType.SEQUENCE_ROI)
        {
            Object roi = sequenceEvent.getSource();
            WritableSheet sheet = getOrCreateSheet(sequence);
            
            switch (sequenceEvent.getType())
            {
            case ADDED:
                update(sequence);
                break;
            
            case CHANGED:
                if (roi != null) update(sequence, (ROI) roi);
                break;
            
            case REMOVED:
                int nbDeleted = (roi == null) ? sheet.getRows() - 1 : 1;
                
                for (int i = nbDeleted; i > 0; i--)
                    sheet.removeRow(i);
                update(sequence);
                break;
            }
            
            // in any case, update the table
            // table.repaint();
            table.updateSheet(sheet);
        }
    }
}
