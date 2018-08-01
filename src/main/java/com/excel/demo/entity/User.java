package com.excel.demo.entity;

import lombok.Data;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import java.io.Serializable;

/**
 * ClassName User
 * Description TODO
 * Author lyk
 * Date 2018/7/17 15:26
 * Version 1.0
 **/
@Data
@Entity
public class User  implements Serializable{
    @Id
    @GeneratedValue
    private Integer id;
    private String name;
    private String phone;
    private String address;
    private String email;

}
