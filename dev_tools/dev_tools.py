import fitz

class DevTools:
    def __init__(self, increment=5, max_search_distance=100):
        self.increment = increment
        self.max_search_distance = max_search_distance

    def find_text_below(self, page, pos):
        for offset in range(0, self.max_search_distance, self.increment):
            rect = fitz.Rect(pos.x0 - 10, pos.y0 + 10, pos.x1 + 20, pos.y1 + offset + self.increment)
            text = page.get_text("text", clip=rect).strip()
            if text:
                return text, rect, offset
        return "", None, None

    def find_text_above(self, page, pos):
        for offset in range(0, self.max_search_distance, self.increment):
            rect = fitz.Rect(pos.x0 - 10, pos.y0 - offset - self.increment, pos.x1  + 20, pos.y0 +10)
            text = page.get_text("text", clip=rect).strip()
            if text:
                return text, rect, offset
        return "", None, None

    def find_text_right(self, page, pos):
        for offset in range(0, self.max_search_distance, self.increment):
            rect = fitz.Rect(pos.x1, pos.y0, pos.x1 + offset + self.increment, pos.y1)
            text = page.get_text("text", clip=rect).strip()
            if text:
                return text, rect, offset
        return "", None, None

    def find_text_left(self, page, pos):
        for offset in range(0, self.max_search_distance, self.increment):
            rect = fitz.Rect(pos.x0 - offset - self.increment - 50, pos.y0, pos.x1, pos.y1)
            text = page.get_text("text", clip=rect).strip()
            if text:
                return text, rect, offset
        return "", None, None

    def find_direction(self, page, pos, direction):
        if direction == "below":
            return self.find_text_below(page, pos)
        elif direction == "above":
            return self.find_text_above(page, pos)
        elif direction == "right":
            return self.find_text_right(page, pos)
        elif direction == "left":
            return self.find_text_left(page, pos)
        else:
            print(f"Unknown direction: {direction}")
            return "", None, None
