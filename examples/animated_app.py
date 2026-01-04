import pygame
import sys
import time

def main():
    pygame.init()
    
    width, height = 640, 480
    screen = pygame.display.set_mode((width, height))
    pygame.display.set_caption("Window Capture Test App")
    
    # Colors
    BLACK = (0, 0, 0)
    RED = (255, 0, 0)
    
    # Ball properties
    ball_pos = [width // 2, height // 2]
    ball_radius = 20
    ball_velocity = [5, 5]
    
    clock = pygame.time.Clock()
    
    running = True
    frame_count = 0
    
    print("Animation started...")
    
    while running:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
        
        # Update ball position
        ball_pos[0] += ball_velocity[0]
        ball_pos[1] += ball_velocity[1]
        
        # Bounce off walls
        if ball_pos[0] <= ball_radius or ball_pos[0] >= width - ball_radius:
            ball_velocity[0] = -ball_velocity[0]
        if ball_pos[1] <= ball_radius or ball_pos[1] >= height - ball_radius:
            ball_velocity[1] = -ball_velocity[1]
            
        # Draw
        screen.fill(BLACK)
        
        # Draw frame counter to verify updates
        font = pygame.font.SysFont(None, 36)
        img = font.render(f'Frame: {frame_count}', True, (255, 255, 255))
        screen.blit(img, (20, 20))
        
        pygame.draw.circle(screen, RED, (int(ball_pos[0]), int(ball_pos[1])), ball_radius)
        
        pygame.display.flip()
        
        frame_count += 1
        clock.tick(60) # Limit to 60 FPS

    pygame.quit()
    sys.exit()

if __name__ == "__main__":
    main()
