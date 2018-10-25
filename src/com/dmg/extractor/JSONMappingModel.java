package com.dmg.extractor;

import com.google.gson.annotations.Expose;
import com.google.gson.annotations.SerializedName;

import java.util.List;

public class JSONMappingModel {

    @SerializedName("targetFilePath")
    @Expose
    private String targetFilePath;
    @SerializedName("sourceFolder")
    @Expose
    private String sourceFolder;
    @SerializedName("all_excel_mappings")
    @Expose
    private List<AllExcelMapping> allExcelMappings = null;

    public String getTargetFilePath() {
        return targetFilePath;
    }

    public void setTargetFilePath(String targetFilePath) {
        this.targetFilePath = targetFilePath;
    }

    public String getSourceFolder() {
        return sourceFolder;
    }

    public void setSourceFolder(String sourceFolder) {
        this.sourceFolder = sourceFolder;
    }

    public List<AllExcelMapping> getAllExcelMappings() {
        return allExcelMappings;
    }

    public void setAllExcelMappings(List<AllExcelMapping> allExcelMappings) {
        this.allExcelMappings = allExcelMappings;
    }


}
