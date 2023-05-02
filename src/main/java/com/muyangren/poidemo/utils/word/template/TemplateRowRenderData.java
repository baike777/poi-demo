package com.muyangren.poidemo.utils.word.template;

import com.deepoove.poi.data.*;
import com.deepoove.poi.data.style.*;
import com.deepoove.poi.policy.TableRenderPolicy;
import com.muyangren.poidemo.entity.FamilyMember;
import com.muyangren.poidemo.entity.Going;
import lombok.Data;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @author: muyangren
 * @Date: 2023/4/27
 * @Description: 将实体类转化为一个表格 因为渲染只接受RowRenderData类型 {@link TableRenderPolicy.Helper#renderRow(XWPFTableRow, RowRenderData)} ()}
 * ps:新版真的难用
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

    private  RowStyle rowStyle;

    public TemplateRowRenderData(List<FamilyMember> familyMemberList, List<Going> goingList) {
        // 初始化样式
        initStyle();
        // 初始化动态数据
        initData(familyMemberList, goingList);
    }

    private void initStyle() {
        // 此处定义样式的优先级高:想看样式获取顺序可以看【TemplateTableRenderPolicy】中的【TableRenderPolicy.Helper.renderRow(xwpfTable.getRow(familyMemberRow), familyRowRenderDataList.get(i))】 谢谢你作者 很规范 nmd

        // 字体样式
        Style style = new Style("宋体", 10);

        // 段落样式
        ParagraphStyle paragraphStyle = new ParagraphStyle();
        paragraphStyle.setDefaultTextStyle(style);
        // ps:这里才是字体居中对齐
        paragraphStyle.setAlign(ParagraphAlignment.CENTER);

        // 表格样式
        CellStyle cellStyle = new CellStyle();
        // ps:表格也需要居中，否则字体不在正中间，会偏上
        cellStyle.setVertAlign(XWPFTableCell.XWPFVertAlign.CENTER);
        cellStyle.setDefaultParagraphStyle(paragraphStyle);

        // 行样式
        this.rowStyle = new RowStyle();
        rowStyle.setDefaultCellStyle(cellStyle);
    }

    private void initData(List<FamilyMember> familyMemberList, List<Going> goingList) {
        List<RowRenderData> newFamilyRowRenderDataList = new ArrayList<>();
        List<RowRenderData> newGoingRowRenderDataList = new ArrayList<>();

        // 判空-家庭人员
        if (CollectionUtils.isNotEmpty(familyMemberList)) {
            for (FamilyMember familyMember : familyMemberList) {
                // 创建一行五表格(根据实际情况来哈)
                List<CellRenderData> cellDataList = new ArrayList<>();
                // 留两个空格(为什么要留两个空白格，大家可以试试自己去添加下行就知道了) 如图：resources/pictures/images3.jpg|images2.jpg
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText("")));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText("")));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(familyMember.getRelationship())));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(familyMember.getName())));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(familyMember.getPosition())));
                RowRenderData rowRenderData = new RowRenderData();
                // 样式
                rowRenderData.setRowStyle(rowStyle);
                rowRenderData.setCells(cellDataList);
                newFamilyRowRenderDataList.add(rowRenderData);
            }
            this.familyRowRenderDataList = newFamilyRowRenderDataList;
        }else {
            // 要是不存在传null值的话，插件里就要多一层判断了
            this.familyRowRenderDataList= Collections.emptyList();
        }

        // 判空-工作情况
        if (CollectionUtils.isNotEmpty(goingList)) {
            for (Going going : goingList) {
                // 创建一行五表格(根据实际情况来哈)
                List<CellRenderData> cellDataList = new ArrayList<>();
                // 保留一个空格(如【家庭成员所示】)
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText("")));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(going.getCompany())));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(going.getAddress())));
                cellDataList.add(new CellRenderData().addParagraph(new ParagraphRenderData().addText(going.getPhone())));
                RowRenderData rowRenderData = new RowRenderData();
                // 样式
                rowRenderData.setRowStyle(rowStyle);
                rowRenderData.setCells(cellDataList);
                newGoingRowRenderDataList.add(rowRenderData);
            }
            this.goingRowRenderDataList = newGoingRowRenderDataList;
        }else {
            this.goingRowRenderDataList= Collections.emptyList();
        }
    }
}
