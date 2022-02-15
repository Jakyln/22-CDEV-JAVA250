package com.example.demo.service;

import com.example.demo.dto.ClientDto;
import com.example.demo.dto.FactureDto;

import java.util.List;

public interface FactureService {
    List<FactureDto> findAll();

}
