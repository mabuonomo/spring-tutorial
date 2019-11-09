package it.squaloxp.redis.cache;

import java.io.Serializable;

public class Post implements Serializable {

    private static final long serialVersionUID = 1L;
    
    private String id;
    private String title;
    private String description;
    private String image;
    private int shares;

    public Post() {
    }

    public Post(String id, String title, String description, String image, int shares) {
        this.id = id;
        this.title = title;
        this.description = description;
        this.image = image;
        this.shares = shares;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public String getImage() {
        return image;
    }

    public void setImage(String image) {
        this.image = image;
    }

    public int getShares() {
        return shares;
    }

    public void setShares(int shares) {
        this.shares = shares;
    }
}
