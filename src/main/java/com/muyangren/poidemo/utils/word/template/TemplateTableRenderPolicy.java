package com.muyangren.poidemo.utils.word.template;

import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.policy.DynamicTableRenderPolicy;
import com.deepoove.poi.policy.TableRenderPolicy;
import com.deepoove.poi.util.TableTools;
import lombok.NoArgsConstructor;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.List;


/**
 * @author: muyangren
 * @Date: 2023/4/27
 * @Description: 自定义渲染插件-for循环【家庭人员】-【工作情况】
 * 这里因为需要操作表格-所以集成DynamicTableRenderPolicy：动态表格插件，允许直接操作表格对象
 * 详情看：http://deepoove.com/poi-tl/#_%E9%BB%98%E8%AE%A4%E6%8F%92%E4%BB%B6
 * @Version: 1.0
 */
@NoArgsConstructor
public class TemplateTableRenderPolicy extends DynamicTableRenderPolicy {

    @Override
    public void render(XWPFTable xwpfTable, Object data) throws Exception {
        if (null == data) {
            return;
        }
        // 获取动态动态数据:为什么可以拿到呢，是因为使用Config插件绑定自定义的渲染规则
        TemplateRowRenderData templateRowRenderData = (TemplateRowRenderData) data;
        // 获取家庭成员
        List<RowRenderData> familyRowRenderDataList = templateRowRenderData.getFamilyRowRenderDataList();
        // 1)判断家庭人员信息是否为空
        if (CollectionUtils.isNotEmpty(familyRowRenderDataList)) {
            // 2)不为空，则有多少条就添加多少行
            // 家庭成员所在行数
             int familyMemberRow = 9;
            // 删除时从下标开始的、所以会删除到空白行
            xwpfTable.removeRow(familyMemberRow);
            // 获取该表家庭成员表头高度
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(familyMemberRow-1);
            // 5)循环插入行
            for (int i = familyRowRenderDataList.size() - 1; i > -1; i--) {
                XWPFTableRow insertNewTableRow = xwpfTable.insertNewTableRow(familyMemberRow);
                insertNewTableRow.setHeight(xwpfTableRow.getHeight());
                // 表格突出就减、缩进去就加：先满足这里在去弄合并的
                for (int j = 0; j < 11; j++) {
                    insertNewTableRow.createCell();
                }
                // (这里就要凭感觉去合并 注意：眼见不一定所得，因为两个表格合并的未必就比一个表格大) 但是
                // 合并单元格 fromCol至toCol 且 fromCol(起始) < toCol(结束), 指的是上一个合并完后的表格数 (基本信息占1格,家庭成员占1个,关系、姓名、职位各占3 刚好11格)
                TableTools.mergeCellsHorizonal(xwpfTable, familyMemberRow, 2, 4); // fromCol开始合并到toCol表格 【关系】下标为2
                TableTools.mergeCellsHorizonal(xwpfTable, familyMemberRow, 3, 5); // fromCol开始合并到toCol表格  【姓名】下标为3
                TableTools.mergeCellsHorizonal(xwpfTable, familyMemberRow, 4, 6); // fromCol开始合并到toCol表格  注：是第一个合并完后的表格数  【职位】下标为4
                // 渲染数据(表格对齐后在进行渲染数据)
                TableRenderPolicy.Helper.renderRow(xwpfTable.getRow(familyMemberRow), familyRowRenderDataList.get(i));
            }
            // 跨行合并行 例如-【基本情况】-【家庭成员】
            // 参数分别为：xwpfTable,需要合并的表格下表,也就【家庭情况】下标为1, 开始合并行，结束行(开始合并+动态数据量)
            int i = familyMemberRow - 1; //减1是因为下标时候从0开始数
            TableTools.mergeCellsVertically(xwpfTable,1,i,familyRowRenderDataList.size()+i);
            TableTools.mergeCellsVertically(xwpfTable,0,0,familyRowRenderDataList.size()+i);
        }
//---------------------------------------以上为家庭成员的数据渲染------------------------------------------------------------

        // 2）获取工作情况
        List<RowRenderData> goingRowRenderDataList = templateRowRenderData.getGoingRowRenderDataList();
        if (CollectionUtils.isNotEmpty(goingRowRenderDataList)) {
            // 2)不为空，则有多少条就添加多少行
            // 工作情况所在行数
            int goingMemberRow = 11;
            if (CollectionUtils.isEmpty(familyRowRenderDataList)) {
            } else {
                goingMemberRow = 11 + familyRowRenderDataList.size() - 1;
            }

            // 删除时从下标开始的、所以会删除到空白行
            xwpfTable.removeRow(goingMemberRow);
            // 获取该表家庭成员表头高度
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(goingMemberRow-1);
            // 5)循环插入行
            for (int i = goingRowRenderDataList.size() - 1; i > -1; i--) {
                XWPFTableRow insertNewTableRow = xwpfTable.insertNewTableRow(goingMemberRow);
                insertNewTableRow.setHeight(xwpfTableRow.getHeight());
                // 表格突出就减、缩进去就加：先满足这里在去弄合并的
                for (int j = 0; j < 11; j++) {
                    insertNewTableRow.createCell();
                }
                // (这里就要凭感觉去合并 注意：眼见不一定所得，因为两个表格合并的未必就比一个表格大) 但是
                // 合并单元格 fromCol至toCol 且 fromCol(起始) < toCol(结束), 指的是上一个合并完后的表格数 (基本信息占1格,家庭成员占1个,关系、姓名、职位各占3 刚好11格)
                TableTools.mergeCellsHorizonal(xwpfTable, goingMemberRow, 1, 2); // fromCol开始合并到toCol表格 【关系】下标为2
                TableTools.mergeCellsHorizonal(xwpfTable, goingMemberRow, 2, 7); // fromCol开始合并到toCol表格  【姓名】下标为3
                TableTools.mergeCellsHorizonal(xwpfTable, goingMemberRow, 3, 4); // fromCol开始合并到toCol表格  注：是第一个合并完后的表格数  【职位】下标为4
                // 渲染数据(表格对齐后在进行渲染数据)
                TableRenderPolicy.Helper.renderRow(xwpfTable.getRow(goingMemberRow), goingRowRenderDataList.get(i));
            }
             //跨行合并行 例如-【基本情况】-【家庭成员】
             //参数分别为：xwpfTable,需要合并的表格下表,也就【家庭情况】下标为1, 开始合并行，结束行(开始合并+动态数据量)
            TableTools.mergeCellsVertically(xwpfTable,0,goingMemberRow-1,goingMemberRow +goingRowRenderDataList.size()-1);
        }
    }
}
