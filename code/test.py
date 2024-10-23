import tkinter as tk
from tkinter import messagebox


class SentenceComparer:
    def __init__(self, reference_sentence, comparison_sentences=None):
        self.reference_sentence = reference_sentence
        if comparison_sentences is None:
            self.comparison_sentences = ["The sky is blue"]
        else:
            self.comparison_sentences = comparison_sentences
            
        self.current_index = 0
        self.total_sentences = len(self.comparison_sentences)
        self.user_choice = None
        self.final_index = None
        self.window = None

    def create_window(self):
        # Create the main window
        self.window = tk.Tk()
        self.window.title("Sentence Comparison")
        self.window.geometry("1000x600")

        # Variable to store the user's answer
        self.choice_var = tk.StringVar()
        self.choice_var.set("Different")  # Default to Different

        # Create frame for progress indicator
        self.progress_label = tk.Label(
            self.window,
            text=f"Comparison {self.current_index + 1} of {self.total_sentences}",
            font=("Arial", 12, "bold")
        )
        self.progress_label.pack(pady=(20, 5))

        # Reference sentence display
        tk.Label(
            self.window,
            text="Reference Sentence:",
            font=("Arial", 12, "bold")
        ).pack(pady=(20, 5))

        tk.Label(
            self.window,
            text=self.reference_sentence,
            font=("Arial", 12),
            wraplength=400
        ).pack(pady=(0, 20))

        # Comparison sentence display
        self.comparison_label = tk.Label(
            self.window,
            text="Comparison Sentence:",
            font=("Arial", 12, "bold")
        )
        self.comparison_label.pack(pady=(20, 5))

        self.comparison_text = tk.Label(
            self.window,
            text=self.comparison_sentences[self.current_index],
            font=("Arial", 12),
            wraplength=400
        )
        self.comparison_text.pack(pady=(0, 20))

        # Create radio buttons for choices
        tk.Label(
            self.window,
            text="Are these sentences the same?",
            font=("Arial", 12, "bold")
        ).pack(pady=(20, 10))

        tk.Radiobutton(
            self.window,
            text="Same",
            variable=self.choice_var,
            value="Same",
            font=("Arial", 12)
        ).pack(pady=5)

        tk.Radiobutton(
            self.window,
            text="Different",
            variable=self.choice_var,
            value="Different",
            font=("Arial", 12)
        ).pack(pady=5)

        # Create navigation buttons frame
        button_frame = tk.Frame(self.window)
        button_frame.pack(pady=20)

        # Previous button
        self.prev_button = tk.Button(
            button_frame,
            text="Previous",
            command=self._prev_pair,
            font=("Arial", 12),
            state=tk.DISABLED if self.current_index == 0 else tk.NORMAL
        )
        self.prev_button.pack(side=tk.LEFT, padx=5)

        # Next button
        self.next_button = tk.Button(
            button_frame,
            text="Next",
            command=self._next_pair,
            font=("Arial", 12),
            state=tk.DISABLED if self.current_index == self.total_sentences - 1 else tk.NORMAL
        )
        self.next_button.pack(side=tk.LEFT, padx=5)

        # Submit button (centered below navigation buttons)
        self.submit_button = tk.Button(
            self.window,
            text="Submit Final Answer",
            command=self._on_submit,
            font=("Arial", 12, "bold"),
            bg="#4CAF50",  # Green background
            fg="white"     # White text
        )
        self.submit_button.pack(pady=20)

    def _update_display(self):
        # Update progress indicator
        self.progress_label.config(text=f"Comparison {self.current_index + 1} of {self.total_sentences}")
        
        # Update comparison sentence
        self.comparison_text.config(text=self.comparison_sentences[self.current_index])
        
        # Update button states
        self.prev_button.config(state=tk.NORMAL if self.current_index > 0 else tk.DISABLED)
        self.next_button.config(state=tk.NORMAL if self.current_index < self.total_sentences - 1 else tk.DISABLED)

    def _prev_pair(self):
        if self.current_index > 0:
            self.current_index -= 1
            self._update_display()

    def _next_pair(self):
        if self.current_index < self.total_sentences - 1:
            self.current_index += 1
            self._update_display()

    def _on_submit(self):
        self.user_choice = self.choice_var.get()
        self.final_index = self.current_index
        self.window.quit()
        self.window.destroy()

    def get_result(self):
        self.create_window()
        self.window.mainloop()
        return self.user_choice, self.final_index


# Example usage
if __name__ == "__main__":
    reference = "The quick brown fox jumps over the lazy dog."
    comparisons = [
        "The quick brown fox jumps over the lazy dog!",
        "The quick brown fox jumps over the lazy dogs.",
        "The quick brown fox jumped over the lazy dog.",
    ]
    
    comparer = SentenceComparer(reference, comparisons)
    result, index = comparer.get_result()
    print(f"User selected '{result}' at comparison {index + 1}")