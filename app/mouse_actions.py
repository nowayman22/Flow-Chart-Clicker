import pyautogui
import random
import math

class MouseActionsMixin:
    def execute_move(self, pos):
        offset = self.loc_offset_variance.get()
        rand_x, rand_y = pos[0] + random.randint(-offset, offset), pos[1] + random.randint(-offset, offset)
        
        speed = 0
        move_mode = self.mouse_move_mode.get()
        start_x, start_y = pyautogui.position()
        distance = math.hypot(rand_x - start_x, rand_y - start_y)

        if move_mode == 'Dynamic':
            screen_w, screen_h = pyautogui.size()
            max_dist = math.hypot(screen_w, screen_h)
            
            min_time = self.min_move_time.get()
            max_time = self.max_move_time.get()
            
            # Linearly interpolate the base speed based on distance
            if max_dist > 0:
                base_speed = min_time + (max_time - min_time) * (distance / max_dist)
            else:
                base_speed = min_time
            speed = max(0, base_speed + random.uniform(-self.speed_variance.get(), self.speed_variance.get()))

        elif move_mode == 'Pixels Per Second':
            pps = self.pixels_per_second.get()
            if pps > 0:
                base_speed = distance / pps
            else:
                base_speed = 0.1 # A small default to prevent instant moves
            speed = max(0, base_speed + random.uniform(-self.speed_variance.get(), self.speed_variance.get()))

        else: # Default to 'Regular' mode
            speed = max(0, self.mouse_speed.get() + random.uniform(-self.speed_variance.get(), self.speed_variance.get()))
            
        pyautogui.moveTo(rand_x, rand_y, duration=speed, tween=pyautogui.easeOutQuad)

    def execute_varied_click(self,pos):
        self.execute_move(pos)
        # Add a check to ensure the click doesn't happen if the move was interrupted
        if not self.running:
            return
        hold = max(0.01, self.hold_duration.get() + random.uniform(-self.hold_duration_variance.get(), self.hold_duration_variance.get()))
        pyautogui.click(duration=hold)

    def execute_action_on_pos(self, action, pos):
        if action == 'Click Object' or action == 'Left Click':
            self.execute_varied_click(pos)
            # Check running state before logging to avoid extraneous logs after stopping
            if self.running:
                self.log_execution(f"Step {self.current_step_index + 1}: Left Clicked near {pos} (Speed: ~{self.mouse_speed.get()}s, Hold: ~{self.hold_duration.get()}s).")
        elif action == 'Click Only':
            if not self.running: return
            hold = max(0.01, self.hold_duration.get() + random.uniform(-self.hold_duration_variance.get(), self.hold_duration_variance.get()))
            pyautogui.click(duration=hold)
            if self.running:
                self.log_execution(f"Step {self.current_step_index + 1}: Clicked at current mouse position.")
        elif action == 'Right Click':
            self.execute_move(pos)
            if not self.running: return # Stop before the click
            pyautogui.rightClick()
            if self.running:
                self.log_execution(f"Step {self.current_step_index + 1}: Right Clicked near {pos}.")
        elif action == 'Move Only':
            self.execute_move(pos)
            if self.running:
                self.log_execution(f"Step {self.current_step_index + 1}: Moved mouse near {pos} (Speed: ~{self.mouse_speed.get()}s).")
   

