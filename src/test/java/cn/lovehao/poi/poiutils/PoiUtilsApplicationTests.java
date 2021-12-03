package cn.lovehao.poi.poiutils;

import cn.lovehao.poi.poiutils.excel.entity.TestExcel;
import cn.lovehao.poi.poiutils.excel.utils.PoiReadUtils;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import java.util.*;


@SpringBootTest
class PoiUtilsApplicationTests {

    @Test
    void contextLoads() {
        List<TestExcel> testExcels =  PoiReadUtils.read("E:\\Users\\84770\\Desktop\\100个商品.xls",TestExcel.class);
        System.out.println(testExcels);
    }

}
