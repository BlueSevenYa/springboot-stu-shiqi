package com.drools;

import com.drools.model.Product;
import org.kie.api.io.Resource;
import org.kie.api.io.ResourceType;
import org.kie.internal.KnowledgeBase;
import org.kie.internal.KnowledgeBaseFactory;
import org.kie.internal.builder.KnowledgeBuilder;
import org.kie.internal.builder.KnowledgeBuilderFactory;
import org.kie.internal.definition.KnowledgePackage;
import org.kie.internal.io.ResourceFactory;
import org.kie.internal.runtime.StatefulKnowledgeSession;

import java.util.Collection;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-09-29-22:14
 */

public class DroolsTest1 {

    public static void main(String[] args) {
        DroolsTest1 droolsTest1=new DroolsTest1();
        droolsTest1.oldExecuteDrools();
    }

    private void oldExecuteDrools() {

        KnowledgeBuilder kbuilder = KnowledgeBuilderFactory.newKnowledgeBuilder();
        Resource ruleFile = ResourceFactory.newFileResource("classpath:com/rules/Rules.drl");
        kbuilder.add(ruleFile,ResourceType.DRL);
        /*kbuilder.add(ResourceFactory.newClassPathResource("com/rules/Rules.drl",
                this.getClass()), ResourceType.DRL);*/
        if (kbuilder.hasErrors()) {
            System.out.println(kbuilder.getErrors().toString());
        }

        Collection<KnowledgePackage> pkgs = kbuilder.getKnowledgePackages();
        // add the package to a rulebase
        KnowledgeBase kbase = KnowledgeBaseFactory.newKnowledgeBase();
        // 将KnowledgePackage集合添加到KnowledgeBase当中
        kbase.addKnowledgePackages(pkgs);

        StatefulKnowledgeSession ksession = kbase.newStatefulKnowledgeSession();
        Product product = new Product();
        product.setType(Product.GOLD);
        ksession.insert(product);
        ksession.fireAllRules();
        ksession.dispose();

        System.out.println("The discount for the product " + product.getType()
                + " is " + product.getDiscount()+"%");
    }


}
