package com.lwb.excel.export.demo;

/**
 * @author liuweibo
 * @date 2019/8/14
 */
public class ParamVar {

    public static void print(String...str) {
        for(String s:str) {
            System.out.println(s);
        }
    }
    public static void main(String[] args) {
        print();
    }
}
