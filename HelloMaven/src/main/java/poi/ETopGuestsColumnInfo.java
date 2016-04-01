package poi;

import org.apache.poi.ss.usermodel.Cell;

/**
 * Created by Alvin on 16/4/1.
 */
public enum  ETopGuestsColumnInfo {

        RANK(0, Cell.CELL_TYPE_STRING, 10 * 512), GUEST_NAME(1, Cell.CELL_TYPE_STRING, 14 * 512), LOGIN_NAME(2, Cell.CELL_TYPE_STRING, 14 * 512), TOTAL(
                3, Cell.CELL_TYPE_STRING, 16 * 512);

        private int	index;
        private int	cellType;
        private int	width;

        private ETopGuestsColumnInfo(int index, int cellType, int width)
        {
            this.index = index;
            this.cellType = cellType;
            this.width = width;
        }

    public int index()
    {
        return index;
    }

    public int cellType()
    {
        return cellType;
    }

    public int width()
    {
        return width;
    }
}

