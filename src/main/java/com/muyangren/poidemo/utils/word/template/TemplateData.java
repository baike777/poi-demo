package com.muyangren.poidemo.utils.word.template;

import com.deepoove.poi.expression.Name;
import com.muyangren.poidemo.entity.Templates;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;


/**
 * @author: muyangren
 * @Date: 2023/4/27
 * @Description: 基础数据+动态数据即{@link TemplateRowRenderData}
 * @Version: 1.0
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class TemplateData {

    @Name("templates")
    private Templates templates;

    @Name("templateRowRenderData")
    private TemplateRowRenderData templateRowRenderData;

}
