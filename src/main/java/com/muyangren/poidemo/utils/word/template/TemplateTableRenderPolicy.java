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
 * 详情请看：http://deepoove.com/poi-tl/#_%E9%BB%98%E8%AE%A4%E6%8F%92%E4%BB%B6
 * @Version: 1.0
 */
@NoArgsConstructor
public class TemplateTableRenderPolicy extends DynamicTableRenderPolicy {

    @Override
    public void render(XWPFTable xwpfTable, Object data) throws Exception {
        if (null == data) {
            return;
        }
     //-----------------------------------先别急,一行一行看下去-------------------------------------------------------
        // 因为我们前面用Config绑定了【List<RowRenderData>】渲染规则
        TemplateRowRenderData templateRowRenderData = (TemplateRowRenderData) data;

        // 一、处理家庭成员数据
        List<RowRenderData> familyRowRenderDataList = templateRowRenderData.getFamilyRowRenderDataList();
        if (CollectionUtils.isNotEmpty(familyRowRenderDataList)) {
            // 1)计算家庭成员表头所在行数 如图：resource/pictures/images4.jpg
             int familyMemberRow = 9;
            // 2)删掉空白内容：xwpfTable.removeRow是从下表0开始计算的，所以这里删除的是空白内容如图：resource/pictures/images5.jpg
            xwpfTable.removeRow(familyMemberRow);
            // 3)获取该表家庭成员表头高度 【familyMemberRow-1】即家庭成员表头的下标  作用：统一高度
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(familyMemberRow-1);
            // 4)循环插入行 ps：这里是一直在第9行插入表格。
            for (int i = familyRowRenderDataList.size() - 1; i > -1; i--) {
                // 4.1)插入表格
                XWPFTableRow insertNewTableRow = xwpfTable.insertNewTableRow(familyMemberRow);
                // 4.1.1)控制表格高度
                insertNewTableRow.setHeight(xwpfTableRow.getHeight());
                // 4.2)每行添加11个表格,这里需要根据实际情况填充。秉持多则退少补原则
                // 如图：假如我添加12个 resource/pictures/images6.jpg
                // 如图：假如我添加10个 resource/pictures/images7.jpg
                // 注：想看效果时记得注释下面的代码 【TableTools.mergeCellsHorizonal】,【TableTools.mergeCellsVertically】
                for (int j = 0; j < 11; j++) {
                    insertNewTableRow.createCell();
                }

                // 基本信息占1格,家庭成员占1个,关系、姓名、职位各占3 刚好11格
                // 5)合并上面创建的11个单元格  fromCol-toCol 且 fromCol(起始) < toCol(结束)
                // 5.1)【关系】下标为2，所以2(fromCol)合并至4(toCol)的单元格
                TableTools.mergeCellsHorizonal(xwpfTable, familyMemberRow, 2, 4);
                // 5.2)【姓名】下标为3，所以3(fromCol)合并至5(toCol)的单元格：注意：按上一个合并后的结果再数格子并合并
                TableTools.mergeCellsHorizonal(xwpfTable, familyMemberRow, 3, 5);
                // 5.3)【职位】下标为4，所以4(fromCol)合并至6(toCol)的单元格：注意：按上一个合并后的结果再数格子并合并
                TableTools.mergeCellsHorizonal(xwpfTable, familyMemberRow, 4, 6);

                // 6) 渲染数据(表格对齐后在进行渲染数据)
                TableRenderPolicy.Helper.renderRow(xwpfTable.getRow(familyMemberRow), familyRowRenderDataList.get(i));
            }
            // 7)跨行合并行
            // 7.1) 合并前：resource/pictures/images8.jpg
            // 7.2) 合并后：resource/pictures/images9.jpg
            int fromRow = familyMemberRow - 1;
            TableTools.mergeCellsVertically(xwpfTable, 1, fromRow, familyRowRenderDataList.size() + fromRow);
            TableTools.mergeCellsVertically(xwpfTable, 0, 0, familyRowRenderDataList.size() + fromRow);
        }
//---------------------------------------以上为家庭成员的数据渲染------------------------------------------------------------

        // 二、处理家庭成员数据
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
