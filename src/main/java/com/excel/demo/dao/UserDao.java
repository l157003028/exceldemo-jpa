package com.excel.demo.dao;

import com.excel.demo.entity.User;
import org.springframework.data.jpa.repository.JpaRepository;

/**
 * InterfaceName UserDao
 * Description TODO
 * Author lyk
 * Date 2018/7/18 13:18
 * Version 1.0
 **/
public interface UserDao extends JpaRepository<User,Integer> {


}
