package com.weixf.render;

import com.deepoove.poi.data.NumberingRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.weixf.example.ExperienceData;

import java.util.List;

public class ResumeDataV2 {

    private PictureRenderData portrait;
    private String name;
    private String job;
    private String sex;
    private String phone;
    private String birthday;
    private String province;
    private String email;
    private String address;
    private String english;
    private String University;
    private String rank;
    private String education;
    private String profession;
    private NumberingRenderData stack;
    private String hobbies;
    private List<ExperienceData> experiences;

    public PictureRenderData getPortrait() {
        return this.portrait;
    }

    public void setPortrait(PictureRenderData portrait) {
        this.portrait = portrait;
    }

    public String getName() {
        return this.name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getJob() {
        return this.job;
    }

    public void setJob(String job) {
        this.job = job;
    }

    public String getSex() {
        return this.sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getPhone() {
        return this.phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getBirthday() {
        return this.birthday;
    }

    public void setBirthday(String birthday) {
        this.birthday = birthday;
    }

    public String getProvince() {
        return this.province;
    }

    public void setProvince(String province) {
        this.province = province;
    }

    public String getEmail() {
        return this.email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getAddress() {
        return this.address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getEnglish() {
        return this.english;
    }

    public void setEnglish(String english) {
        this.english = english;
    }

    public String getUniversity() {
        return this.University;
    }

    public void setUniversity(String University) {
        this.University = University;
    }

    public String getRank() {
        return this.rank;
    }

    public void setRank(String rank) {
        this.rank = rank;
    }

    public String getEducation() {
        return this.education;
    }

    public void setEducation(String education) {
        this.education = education;
    }

    public String getProfession() {
        return this.profession;
    }

    public void setProfession(String profession) {
        this.profession = profession;
    }

    public NumberingRenderData getStack() {
        return this.stack;
    }

    public void setStack(NumberingRenderData stack) {
        this.stack = stack;
    }

    public String getHobbies() {
        return this.hobbies;
    }

    public void setHobbies(String hobbies) {
        this.hobbies = hobbies;
    }

    public List<ExperienceData> getExperiences() {
        return experiences;
    }

    public void setExperiences(List<ExperienceData> experiences) {
        this.experiences = experiences;
    }

}
