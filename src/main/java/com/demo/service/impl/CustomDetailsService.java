//package com.demo.service.impl;
//
//import com.demo.entities.MyUser;
//import com.demo.entities.Role;
//import com.demo.repository.UserRepository;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.boot.autoconfigure.security.servlet.UserDetailsServiceAutoConfiguration;
//import org.springframework.security.core.GrantedAuthority;
//import org.springframework.security.core.authority.SimpleGrantedAuthority;
//import org.springframework.security.core.userdetails.UserDetails;
//import org.springframework.security.core.userdetails.UserDetailsService;
//import org.springframework.security.core.userdetails.UsernameNotFoundException;
//import org.springframework.stereotype.Service;
//
//import java.util.ArrayList;
//import java.util.List;
//
//@Service
//public class CustomDetailsService implements UserDetailsService
//{
//    @Autowired
//    private UserRepository userRepository;
//    @Override
//    public UserDetails loadUserByUsername(String s) throws UsernameNotFoundException
//    {
//        MyUser user = userRepository.findByUsername(s);
//        List<GrantedAuthority> authorities = new ArrayList<>();
//        for (Role role : user.getRoles())
//            authorities.add(new SimpleGrantedAuthority("ROLE_"+role.getCode()));
//        MyUser myUser = new MyUser(user.getUsername() , user.getPassword() , true ,
//                true ,true , true , authorities);
//        myUser.setActive(user.getActive());
//        myUser.setName(user.getName());
//        return myUser;
//    }
//}
