package com.muyangren.poidemo;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.muyangren.poidemo.entity.FamilyMember;
import com.muyangren.poidemo.entity.Going;
import com.muyangren.poidemo.entity.Templates;
import com.muyangren.poidemo.utils.word.template.TemplateData;
import com.muyangren.poidemo.utils.word.template.TemplateRowRenderData;
import com.muyangren.poidemo.utils.word.template.TemplateTableRenderPolicy;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

/**
 * ps：如果你是刚入门的，建议就别看了，徒增痛苦，告诉leader这需求不行
 */
@SpringBootTest
class PoiDemoApplicationTests {
    /**
     * 按照我写的顺序先看一遍
     */
    @Test
    void contextLoads()  {
        try {
        // 一、初始化数据
        // 1）初始化模板数据-实际是去数据库查询-自行整合
          Templates templates = initTemplates();
        // 2）初始化家庭成员信息-实际是去数据库查询-自行整合
        List<FamilyMember> familyMemberList = initFamilyMember();
        // 3）初始化工作情况-实际是去数据库查询-自行整合
        List<Going> goingList = initGoing();

        // 二、初始化动态数据-并整合一个完整的对象
        // 1）处理动态数据->转为RowRenderData类型即List<Going>->List<RowRenderData>
        TemplateRowRenderData templateRowRenderData = new TemplateRowRenderData(familyMemberList,goingList);
        // 2）完整数据
        TemplateData templateData = new TemplateData(templates,templateRowRenderData);

        // 三、绑定插件
        // 1）插件绑定-【TemplateTableRenderPolicy】插件中data能获取到【TemplateRowRenderData】动态数据的关键
        // 注意：模板中也是根据该字段进行渲染：templateRowRenderData 如图resources/pictures/images1.jpg   特别注意不要用到了中文画框花，鼠鼠我郁闷了一早上也想不明道插件为什么不生效
        Configure config = Configure.builder().bind("templateRowRenderData", new TemplateTableRenderPolicy()).build();

        //  四、导出
        ClassPathResource classPathResource = new ClassPathResource("templates" + File.separator + "template.docx");
        XWPFTemplate template = XWPFTemplate.compile(classPathResource.getInputStream(),config).render(
                templateData);
        Assertions.assertNotNull(template);
        // 通过浏览器下载自行整合下即可
        // controller获取HttpServerResponse
            template.writeAndClose(Files.newOutputStream(Paths.get("C:\\Users\\WUDI\\Desktop\\export-word.docx")));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    private Templates initTemplates() {
        // 填充一些基本信息
        Templates templates = new Templates();
        templates.setName("牧羊人OVO");
        templates.setAliases("牧羊人OVo");
        templates.setDeptName("青龙小组");
        templates.setSexName("未知");
        templates.setPeoples("汉族");
        templates.setBirth("2000");
        templates.setCulture("小学");
        return templates;
    }

    /**
     * 测试数据、实际是需要查询数据库的
     * @return
     */
    private List<FamilyMember> initFamilyMember() {
        List<FamilyMember> familyMemberList = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            FamilyMember familyMember = new FamilyMember();
            familyMember.setRelationship("关系" + i);
            familyMember.setName("姓名" + i);
            familyMember.setPosition("职位" + i);
            familyMemberList.add(familyMember);
        }
        return familyMemberList;
    }

    /**
     * 测试数据
     *
     * @return
     */
    private List<Going> initGoing() {
        List<Going> goingList = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            Going going = new Going();
            going.setCompany("工作单位" + i);
            going.setAddress("地址" + i);
            going.setPhone("程联系电话" + i);
            goingList.add(going);
        }
        return goingList;
    }
}
