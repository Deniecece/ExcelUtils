package com.excel.service.impl;

import com.excel.domain.BusClick;
import com.excel.service.ExcelService;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelServiceImpl implements ExcelService {

    @Override
    public List<BusClick> getBusClick() {
        List<BusClick> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            BusClick busClick = new BusClick(Integer.toString(i+2), Integer.toString(i+3), Integer.toString(i+4), Integer.toString(i+5), Integer.toString(i+6));
            list.add(busClick);
        }
        return list;
    }
}
