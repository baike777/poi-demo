package com.muyangren.poidemo.utils.word.template;

import com.deepoove.poi.data.*;
import com.deepoove.poi.data.style.CellStyle;
import com.deepoove.poi.policy.TableRenderPolicy;
import com.muyangren.poidemo.entity.FamilyMember;
import com.muyangren.poidemo.entity.Going;
import com.sun.org.apache.bcel.internal.generic.NEW;
import lombok.Data;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.ArrayList;
import java.util.List;

/**
 * @author: muyangren
 * @Date: 2023/4/27
 * @Description: 将实体类转化为一个表格 因为渲染只接受RowRenderData类型 {@link TableRenderPolicy.Helper#renderRow(XWPFTableRow, RowRenderData)}  ()}
 * @Version: 1.0
 */
@Data
public class TemplateRowRenderData {
    /**
     * 家庭人员
     */
    private List<RowRenderData> familyRowRenderDataList;

    /**
     * 工作情况
     */
    private List<RowRenderData> goingRowRenderDataList;

    public TemplateRowRenderData(List<FamilyMember> familyMemberList, List<Going> goingList) {
        // 初始化动态数据
        initData(familyMemberList, goingList);
    }

    private void initData(List<FamilyMember> familyMemberList, List<Going> goingList) {
        List<RowRenderData> newFamilyRowRenderDataList = new ArrayList<>();
        List<RowRenderData> newGoingRowRenderDataList = new ArrayList<>();

        // 判空-家庭人员
        if (CollectionUtils.isNotEmpty(familyMemberList)) {
            for (FamilyMember familyMember : familyMemberList) {
                // 创建一个行数据
                List<CellRenderData> cellDataList = new ArrayList<>();
                // 留两个空格，方便合并
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText("")));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText("")));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(familyMember.getRelationship())));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(familyMember.getName())));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(familyMember.getPosition())));
                RowRenderData rowRenderData = new RowRenderData();
                rowRenderData.setCells(cellDataList);
                newFamilyRowRenderDataList.add(rowRenderData);
            }
            this.familyRowRenderDataList = newFamilyRowRenderDataList;
        }

        // 判空-工作情况
        if (CollectionUtils.isNotEmpty(goingList)) {
            for (Going going : goingList) {
                // 留一个空格，给【工作情况】行合并
                List<CellRenderData> cellDataList = new ArrayList<>();
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText("")));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(going.getCompany())));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(going.getAddress())));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(going.getPhone())));
                RowRenderData rowRenderData = new RowRenderData();
                rowRenderData.setCells(cellDataList);
                newGoingRowRenderDataList.add(rowRenderData);
            }
            this.goingRowRenderDataList = newGoingRowRenderDataList;
        }
    }
}
