using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuildExcel
{
    /**
         * Utility methods used to convert Excel's character based column and row
         * size measurements into pixels and/or millimetres. The class also contains
         * various constants that are required in other calculations.
         *
         * @author xio[darjino@hotmail.com]
         * @version 1.01 30th July 2009.
         *      Added by Mark Beardsley [msb at apache.org].
         *          Additional constants.
         *          widthUnits2Millimetres() and millimetres2Units() methods.
         */

    public static class ConvertImageUnits
    {

        // Each cell conatins a fixed number of co-ordinate points; this number
        // does not vary with row height or column width or with font. These two
        // constants are defined below.
        public static int TOTAL_COLUMN_COORDINATE_POSITIONS = 1023; // MB
        public static int TOTAL_ROW_COORDINATE_POSITIONS = 255; // MB
        // The resoultion of an image can be expressed as a specific number
        // of pixels per inch. Displays and printers differ but 96 pixels per
        // inch is an acceptable standard to beging with.
        public static int PIXELS_PER_INCH = 96; // MB
        // Cnstants that defines how many pixels and points there are in a
        // millimetre. These values are required for the conversion algorithm.
        public static double PIXELS_PER_MILLIMETRES = 3.78; // MB
        public static double POINTS_PER_MILLIMETRE = 2.83; // MB
        // The column width returned by HSSF and the width of a picture when
        // positioned to exactly cover one cell are different by almost exactly
        // 2mm - give or take rounding errors. This constant allows that
        // additional amount to be accounted for when calculating how many
        // celles the image ought to overlie.
        public static double CELL_BORDER_WIDTH_MILLIMETRES = 2.0D; // MB
        public static short EXCEL_COLUMN_WIDTH_FACTOR = 256;
        public static int UNIT_OFFSET_LENGTH = 8;

        public static int[] UNIT_OFFSET_MAP = new int[]
        {0, 36, 73, 109, 146, 182, 219};

        /**
            * pixel units to excel width units(units of 1/256th of a character width)
            * @param pxs
            * @return
            */

        public static short pixel2WidthUnits(int pxs)
        {
            short widthUnits = (short) (EXCEL_COLUMN_WIDTH_FACTOR*
                                        (pxs/UNIT_OFFSET_LENGTH));
            widthUnits += (short) UNIT_OFFSET_MAP[(pxs%UNIT_OFFSET_LENGTH)];
            return widthUnits;
        }

        /**
             * excel width units(units of 1/256th of a character width) to pixel
             * units.
             *
             * @param widthUnits
             * @return
             */

        public static int widthUnits2Pixel(short widthUnits)
        {
            int pixels = (widthUnits/EXCEL_COLUMN_WIDTH_FACTOR)
                         *UNIT_OFFSET_LENGTH;
            int offsetWidthUnits = widthUnits%EXCEL_COLUMN_WIDTH_FACTOR;
            pixels += (int) Math.Round(offsetWidthUnits/
                                       ((float) EXCEL_COLUMN_WIDTH_FACTOR/UNIT_OFFSET_LENGTH));
            return pixels;
        }

        /**
             * Convert Excel's width units into millimetres.
             *
             * @param widthUnits The width of the column or the height of the
             *                   row in Excel's units.
             * @return A primitive double that contains the columns width or rows
             *         height in millimetres.
             */

        public static double widthUnits2Millimetres(short widthUnits)
        {
            return (ConvertImageUnits.widthUnits2Pixel(widthUnits)/
                    ConvertImageUnits.PIXELS_PER_MILLIMETRES);
        }

        /**
             * Convert into millimetres Excel's width units..
             *
             * @param millimetres A primitive double that contains the columns
             *                    width or rows height in millimetres.
             * @return A primitive int that contains the columns width or rows
             *         height in Excel's units.
             */

        public static int millimetres2WidthUnits(double millimetres)
        {
            return (ConvertImageUnits.pixel2WidthUnits((int) (millimetres*
                                                              ConvertImageUnits.PIXELS_PER_MILLIMETRES)));
        }
    }
}
