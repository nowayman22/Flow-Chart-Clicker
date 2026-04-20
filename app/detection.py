import cv2
import numpy as np
import os
import math

class DetectionMixin:
    def find_png(self, screen_cv, offset, step):
        image_mode = step.get('image_mode', 'Grayscale')
        if image_mode == 'Grayscale': screen_processed = cv2.cvtColor(screen_cv, cv2.COLOR_BGR2GRAY)
        elif image_mode == 'Binary (B&W)': gray = cv2.cvtColor(screen_cv, cv2.COLOR_BGR2GRAY); _, screen_processed = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        else: screen_processed = screen_cv
        
        templates_to_check = []
        if step['mode'] == 'file' and step['path']:
            templates_to_check.append(self.load_template(step['path'], image_mode))
        elif step['mode'] == 'folder' and step['path'] and os.path.isdir(step['path']):
            folder_cache_key = f"{step['path']}|{image_mode}"
            if folder_cache_key not in self.folder_image_cache:
                self.folder_image_cache[folder_cache_key] = []
                image_paths = [os.path.join(step['path'], fname) for fname in os.listdir(step['path']) if fname.lower().endswith('.png')]
                for fpath in image_paths:
                    template_data = self.load_template(fpath, image_mode)
                    if template_data[0] is not None:
                        self.folder_image_cache[folder_cache_key].append(template_data)
            templates_to_check = self.folder_image_cache.get(folder_cache_key, [])

        find_first = step.get('find_first_match', False)
        if not find_first:
            best_match_pos, max_confidence = None, -1

        for template_data in templates_to_check:
            match = self.find_template_in_region(screen_processed, offset, template_data, step['threshold'])
            if match and math.isfinite(match[2]):
                if find_first:
                    return match[0:2], match[2]
                
                if match[2] > max_confidence:
                    max_confidence = match[2]
                    best_match_pos = match[0:2]

        if find_first:
            return None, 0
        else:
            return best_match_pos, max_confidence if max_confidence > -1 else 0

    def find_and_count_png(self, screen_cv, offset, step):
        """
        Finds all occurrences of template(s) in the screen region and returns the count.
        Uses an optimized, built-in OpenCV method to group overlapping matches.
        """
        image_mode = step.get('image_mode', 'Grayscale')
        if image_mode == 'Grayscale':
            screen_processed = cv2.cvtColor(screen_cv, cv2.COLOR_BGR2GRAY)
        elif image_mode == 'Binary (B&W)':
            gray = cv2.cvtColor(screen_cv, cv2.COLOR_BGR2GRAY)
            _, screen_processed = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        else:
            screen_processed = screen_cv
        
        templates_to_check = []
        if step['mode'] == 'file' and step['path']:
            templates_to_check.append(self.load_template(step['path'], image_mode))
        elif step['mode'] == 'folder' and step['path'] and os.path.isdir(step['path']):
            folder_cache_key = f"{step['path']}|{image_mode}"
            if folder_cache_key not in self.folder_image_cache: # Caching logic
                self.folder_image_cache[folder_cache_key] = []
                image_paths = [os.path.join(step['path'], fname) for fname in os.listdir(step['path']) if fname.lower().endswith('.png')]
                for fpath in image_paths:
                    template_data = self.load_template(fpath, image_mode)
                    if template_data[0] is not None: self.folder_image_cache[folder_cache_key].append(template_data)
            templates_to_check = self.folder_image_cache.get(folder_cache_key, [])
        
        all_rects = []
        threshold = step['threshold']

        for template, mask in templates_to_check:
            if template is None: continue
            h, w = template.shape[:2]
            if any(s_dim < t_dim for s_dim, t_dim in zip(screen_processed.shape, template.shape)):
                continue

            locs = None
            if mask is not None:
                # Use TM_SQDIFF_NORMED for transparent PNGs for consistency with single-detect
                res = cv2.matchTemplate(screen_processed, template, cv2.TM_SQDIFF_NORMED, mask=mask)
                # For SQDIFF, a lower score is better, so we invert the threshold check
                locs = np.where(res <= (1.0 - threshold))
            else:
                # Use the standard method for opaque images
                res = cv2.matchTemplate(screen_processed, template, cv2.TM_CCOEFF_NORMED)
                locs = np.where(res >= threshold)
            
            # Create a list of all found rectangles [x, y, width, height]
            for pt in zip(*locs[::-1]):
                all_rects.append([pt[0], pt[1], w, h])

        if not all_rects:
            return 0
        
        # Use OpenCV's optimized groupRectangles function to merge overlapping boxes.
        grouped_rects, _ = cv2.groupRectangles(all_rects, groupThreshold=1, eps=0.2)

        return len(grouped_rects)

    def find_and_count_color(self, screen_cv, offset, step):
        """
        Finds all occurrences of a color in the screen region and returns the count.
        """
        rgb = step['rgb']
        tolerance = step['tolerance']
        color_space = step.get('color_space', 'HSV')
        min_area = step.get('min_pixel_area', 10)
        
        if color_space == 'RGB':
            lower = np.array([max(0, rgb[2]-tolerance), max(0, rgb[1]-tolerance), max(0, rgb[0]-tolerance)]) 
            upper = np.array([min(255, rgb[2]+tolerance), min(255, rgb[1]+tolerance), min(255, rgb[0]+tolerance)])
            mask = cv2.inRange(screen_cv, lower, upper)
        else: # HSV
            hsv = cv2.cvtColor(screen_cv, cv2.COLOR_BGR2HSV)
            target_hsv = cv2.cvtColor(np.uint8([[list(reversed(rgb))]]), cv2.COLOR_BGR2HSV)[0][0]
            h, s, v = int(target_hsv[0]), int(target_hsv[1]), int(target_hsv[2])
            h_tol, s_tol, v_tol = int(tolerance*1.8), int(tolerance*2.5), int(tolerance*2.5)
            lower = np.array([max(0,h-h_tol), max(0,s-s_tol), max(0,v-v_tol)])
            upper = np.array([min(179,h+h_tol), min(255,s+s_tol), min(255,v+v_tol)])
            mask = cv2.inRange(hsv, lower, upper)

        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        valid_contours = [c for c in contours if cv2.contourArea(c) > min_area]
        
        return len(valid_contours)

    def load_template(self, path, image_mode='Grayscale'):
        cache_key = f"{path}|{image_mode}"
        if cache_key not in self.template_cache:
            try:
                img = cv2.imread(path, cv2.IMREAD_UNCHANGED)
                if img is None: raise ValueError("Image not found or unable to read.")
                
                mask = None
                if len(img.shape) == 3 and img.shape[2] == 4:
                    alpha = img[:, :, 3]
                    _, mask = cv2.threshold(alpha, 1, 255, cv2.THRESH_BINARY)
                    img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                
                processed_template = None
                if image_mode == 'Color':
                    if len(img.shape) == 2: 
                        processed_template = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
                    else: 
                        processed_template = img
                else: 
                    if len(img.shape) == 3: 
                        processed_template = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                    else: 
                        processed_template = img

                    if image_mode == 'Binary (B&W)':
                        _, processed_template = cv2.threshold(processed_template, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

                self.template_cache[cache_key] = (processed_template, mask)
            except Exception as e: 
                self.log(f"Error loading template {os.path.basename(path)}: {e}", "red")
                self.template_cache[cache_key] = (None, None)
        return self.template_cache[cache_key]

    def rgb_to_hex(self,rgb): return f'#{int(rgb[0]):02x}{int(rgb[1]):02x}{int(rgb[2]):02x}' if rgb else '#000000'

    def preprocess_for_ocr(self, img_cv):
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY); inverted = cv2.bitwise_not(gray)
        _, thresh = cv2.threshold(inverted, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return thresh

    def find_color_on_screen_hsv(self,img_bgr,offset,step):
        rgb = step.get('rgb', (255,0,0)); tolerance = step.get('tolerance', 2); min_area = step.get('min_pixel_area', 10)
        hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV); target_hsv = cv2.cvtColor(np.uint8([[list(reversed(rgb))]]), cv2.COLOR_BGR2HSV)[0][0]
        h, s, v = int(target_hsv[0]), int(target_hsv[1]), int(target_hsv[2]); h_tol, s_tol, v_tol = int(tolerance*1.8), int(tolerance*2.5), int(tolerance*2.5)
        lower = np.array([max(0,h-h_tol), max(0,s-s_tol), max(0,v-v_tol)]); upper = np.array([min(179,h+h_tol), min(255,s+s_tol), min(255,v+v_tol)])
        mask = cv2.inRange(hsv, lower, upper); contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if contours:
            largest = max(contours, key=cv2.contourArea); area = cv2.contourArea(largest)
            if area > min_area:
                M = cv2.moments(largest)
                if M['m00'] != 0: return (int(M['m10']/M['m00'])+offset[0], int(M['m01']/M['m00'])+offset[1]), area
        return None, 0

    def find_color_on_screen_rgb(self, img_bgr, offset, step):
        rgb = step.get('rgb', (255,0,0)); tolerance = step.get('tolerance', 2); min_area = step.get('min_pixel_area', 10)
        lower = np.array([max(0, rgb[2]-tolerance), max(0, rgb[1]-tolerance), max(0, rgb[0]-tolerance)]) 
        upper = np.array([min(255, rgb[2]+tolerance), min(255, rgb[1]+tolerance), min(255, rgb[0]+tolerance)])
        mask = cv2.inRange(img_bgr, lower, upper)
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if contours:
            largest = max(contours, key=cv2.contourArea); area = cv2.contourArea(largest)
            if area > min_area:
                M = cv2.moments(largest)
                if M['m00'] != 0: return (int(M['m10']/M['m00'])+offset[0], int(M['m01']/M['m00'])+offset[1]), area
        return None, 0

    def find_template_in_region(self, screen_processed, offset, template_data, threshold):
        """
        Finds a template in a pre-processed screen region using an optimized matching method.
        Uses a fast, mask-aware method for transparent PNGs.
        """
        template_processed, mask = template_data
        if template_processed is None or any(s_dim < t_dim for s_dim, t_dim in zip(screen_processed.shape, template_processed.shape)):
            return None

        h, w = template_processed.shape[:2]

        if mask is not None:
            res = cv2.matchTemplate(screen_processed, template_processed, cv2.TM_SQDIFF_NORMED, mask=mask)
            min_val, _, min_loc, _ = cv2.minMaxLoc(res)
            
            confidence = 1.0 - min_val
            if confidence >= threshold:
                return (min_loc[0] + w // 2 + offset[0], min_loc[1] + h // 2 + offset[1], confidence)
        else:
            res = cv2.matchTemplate(screen_processed, template_processed, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            if max_val >= threshold:
                return (max_loc[0] + w // 2 + offset[0], max_loc[1] + h // 2 + offset[1], max_val)
                
        return None


