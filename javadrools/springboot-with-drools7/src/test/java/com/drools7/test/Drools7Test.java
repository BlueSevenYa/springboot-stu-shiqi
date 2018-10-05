package com.drools7.test;

import com.drools7.model.Product;
import org.junit.Test;
import org.kie.api.KieServices;
import org.kie.api.runtime.KieContainer;
import org.kie.api.runtime.KieSession;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-10-03-17:14
 */

public class Drools7Test {

    @Test
    /*
     * 简单使用drools7进行demo测试
     */
    public void test1(){
        System.out.println("test...........");
        KieServices ks=KieServices.Factory.get();
        KieContainer kieContainer=ks.getKieClasspathContainer();
        KieSession kieSession=kieContainer.newKieSession("ksession-rule");

        Product product=new Product();
        product.setType(Product.GOLD);

        kieSession.insert(product);
        int count=kieSession.fireAllRules();
        System.out.println("命中了"+count+"条规则");
        System.out.println("商品" +product.getType() + "的商品折扣为" + product.getDiscount() + "%。");
    }

}
