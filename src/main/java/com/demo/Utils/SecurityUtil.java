//package com.demo.Utils;
//
//import org.springframework.security.core.GrantedAuthority;
//import org.springframework.security.core.context.SecurityContextHolder;
//
//import java.util.ArrayList;
//import java.util.List;
//
//public class SecurityUtil
//{
//    public static List<String> getRoles()
//    {
//        List<GrantedAuthority> authorities = (List<GrantedAuthority>) SecurityContextHolder.getContext().getAuthentication().getAuthorities();
//        List<String> list = new ArrayList<>();
//        for (GrantedAuthority authority : authorities)
//            list.add(authority.getAuthority());
//        return list;
//    }
//}
