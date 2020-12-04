//package com.demo.entities;
//
//import javax.persistence.*;
//import java.util.Set;
//
//@Entity
//@Table(name = "role")
//public class Role extends BaseEntity
//{
//    @Column(name = "code")
//    private String code;
//    @Column(name = "name")
//    private String name;
//    @ManyToMany(mappedBy = "roles")
//    private Set<MyUser> myUsers;
//    @OneToMany(mappedBy = "role" , fetch = FetchType.EAGER)
//    private Set<Url> urls;
//
//    public Set<MyUser> getUsers() {
//        return myUsers;
//    }
//
//    public void setUsers(Set<MyUser> myUsers) {
//        this.myUsers = myUsers;
//    }
//
//    public String getCode() {
//        return code;
//    }
//
//    public void setCode(String code) {
//        this.code = code;
//    }
//
//    public String getName() {
//        return name;
//    }
//
//    public void setName(String name) {
//        this.name = name;
//    }
//}
