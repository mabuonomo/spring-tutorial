package it.squaloxp.redis.cache;

import org.springframework.stereotype.Service;

import it.squaloxp.redis.cache.exception.PostNotFoundException;

import javax.annotation.PostConstruct;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;

@Service
public class PostService {

    private List<Post> posts;

    @PostConstruct
    public void init() {
        posts = new ArrayList<>();
        posts.add(new Post("IDX001", "Cyberpunk is near", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 555));
        posts.add(new Post("IDX002", "Welcome aboard of the hype train", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 55));
        posts.add(new Post("IDX003", "How to improve programming skills ", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 555));
        posts.add(new Post("IDX004", "Top exercises for IT people", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 55));
        posts.add(new Post("IDX005", "Case study of 75 years project", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 55));
        posts.add(new Post("IDX006", "Machine Learning", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 255));
        posts.add(new Post("IDX007", "Memory leaks, how to find them ", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 55));
        posts.add(new Post("IDX008", "Robots builds robots", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 155));
        posts.add(new Post("IDX009", "Quantum algorithms, from the scratch", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 455));
        posts.add(new Post("IDX010", "Coding, coding, coding", "Description",
                "https://random-cdn.posts/images/xvn84934fnls.jpg", 255));
    }

    public void updatePost(Post post) {
        posts = posts.stream().filter(p -> p.getId() != post.getId()).collect(Collectors.toList());
        posts.add(post);
    }

    public void deletePost(String postID) {
        posts = posts.stream().filter(p -> !p.getId().equals(postID)).collect(Collectors.toList());
    }

    /**
     * Load post content from DB *Long running method
     */
    public Post getPostByID(String postID) throws PostNotFoundException {
        try {
            TimeUnit.SECONDS.sleep(1);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        return posts.stream().filter(p -> p.getId().equals(postID)).findFirst()
                .orElseThrow(() -> new PostNotFoundException("Cannot find post with id:" + postID));
    }

    /**
     * Load top ten posts from DB *Long running method
     */
    public List<Post> getTopPosts() {
        try {
            TimeUnit.SECONDS.sleep(2);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        return posts;
    }
}
