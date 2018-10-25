package com.dmg.extractor;

import com.google.gson.annotations.Expose;
import com.google.gson.annotations.SerializedName;

public class CellMapping {

    @SerializedName("source_cell_name")
    @Expose
    private String sourceCellName;
    @SerializedName("source_row_position")
    @Expose
    private String sourceRowPosition;
    @SerializedName("destination_cell_name")
    @Expose
    private String destinationCellName;
    @SerializedName("destination_row_position")
    @Expose
    private String destinationRowPosition;

    @SerializedName("source_search_string")
    @Expose
    private String sourceSearchString;


    public String getSourceCellName() {
        return sourceCellName;
    }

    public void setSourceCellName(String sourceCellName) {
        this.sourceCellName = sourceCellName;
    }

    public String getSourceSearchString() {
        return sourceSearchString;
    }

    public void setSourceSearchString(String sourceSearchString) {
        this.sourceSearchString = sourceSearchString;
    }


    public String getSourceRowPosition() {
        return sourceRowPosition;
    }

    public void setSourceRowPosition(String sourceRowPosition) {
        this.sourceRowPosition = sourceRowPosition;
    }

    public String getDestinationCellName() {
        return destinationCellName;
    }

    public void setDestinationCellName(String destinationCellName) {
        this.destinationCellName = destinationCellName;
    }

    public String getDestinationRowPosition() {
        return destinationRowPosition;
    }

    public void setDestinationRowPosition(String destinationRowPosition) {
        this.destinationRowPosition = destinationRowPosition;
    }


}
