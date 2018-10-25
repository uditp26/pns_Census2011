package com.dmg.extractor;

import com.google.gson.annotations.Expose;
import com.google.gson.annotations.SerializedName;

import java.util.List;

public class AllExcelMapping {

    @SerializedName("excel_name")
    @Expose
    private String excelName;




    @SerializedName("cell_mappings")
    @Expose
    private List<CellMapping> cellMappings = null;

    public String getExcelName() {
        return excelName;
    }

    public void setExcelName(String excelName) {
        this.excelName = excelName;
    }


    public List<CellMapping> getCellMappings() {
        return cellMappings;
    }

    public void setCellMappings(List<CellMapping> cellMappings) {
        this.cellMappings = cellMappings;
    }

    @Override
    public boolean equals(Object o) {
        if(o instanceof AllExcelMapping) {
            AllExcelMapping allExcelMapping = (AllExcelMapping) o;

            if (excelName.startsWith(allExcelMapping.getExcelName())) {
                return true;
            }
        }
        return false;
    }

//    @Override
//    public int hashCode() {
//        return super.hashCode();
//    }
}
