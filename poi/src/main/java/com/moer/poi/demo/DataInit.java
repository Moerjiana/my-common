package com.moer.poi.demo;

import com.moer.poi.po.UserPo;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DataInit {
    public List<UserPo> init(){
        List<UserPo> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserPo po = new UserPo();
            po.setId(i);
            po.setUsername("摩尔"+i);
            po.setTitle("标题"+i);
            po.setSex(String.valueOf((i % 2)));
            po.setStatus((i % 2));
            po.setMoney(String.valueOf(100+i));
            list.add(po);
        }
        return list;
    }

    public Map<Integer,String> initSexMap(){
        Map<Integer,String> map = new HashMap<>();
        map.put(0,"男");
        map.put(1,"女");
        return map;
    }

    public Map<Integer,String> initStatusMap(){
        Map<Integer,String> map = new HashMap<>();
        map.put(0,"启用");
        map.put(1,"禁用");
        return map;
    }

    public static void main(String[] args) {
        for (int i = 0; i < 10; i++) {
            System.out.println(i % 2);
        }
    }
}
