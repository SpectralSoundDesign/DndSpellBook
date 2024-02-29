import tkinter as tk
from tkinter import ttk
import handle_excel as hs
from tkinter import simpledialog
from tkinter import scrolledtext
import json
from ttkthemes import ThemedTk

class SpellbookApp(ThemedTk):
    def __init__(self):
        super().__init__()
        self.set_theme('arc')  # Set the theme to "arc" for a modern look
        self.configure(bg='#161A30')
        self.style = ttk.Style()
        self.style.configure("TLabel", foreground="white", background="#161A30")
        self.style.configure("TLabel1.TLabel", foreground="black", background="#b6bbc4")
        self.style.configure("TButton", foreground="#161A30", background="#161a30")
        self.style.configure("TButton1.TButton", foreground="#161A30", background="#b6bbc4")
        self.style.configure("TEntry", foreground="#161A30", background="#F0ECE5")
        self.style.configure("TListbox", foreground="#161A30", background="#F0ECE5", bordercolor="#161A30")
        self.style.configure("TScrollbar", foreground="#161A30", background="#F0ECE5")
        self.style.configure("TFrame", foreground="#161A30", background="#161A30")
        self.style.configure("Frame1.TFrame", foreground="#b6bbc4", background="#b6bbc4")
        self.style.configure("TNotebook", background="#b6bbc4", foreground="#b6bbc4")
        self.style.configure("TNotebook.Tab", background="#161a30", foreground="#161a30")
        self.title("D&D Spellbook")
        self.geometry("400x300")
        
        self.character_name_label = ttk.Label(self, text="Character: " + "Cygnus", font=("Helvetica", 14, "bold"))
        self.character_name_label.pack(padx=10, pady=10, anchor=tk.NW)
        
        search_spell_frame = ttk.Frame(self)
        search_spell_frame.pack(anchor=tk.NW, pady=10)
        
        self.search_spell_label = ttk.Label(search_spell_frame, text="Search for a spell: ", font=("Helvetica", 14, "bold"))
        self.search_spell_label.pack(side=tk.LEFT, padx=10, pady=10)
        
        self.spell_search_var = tk.StringVar()
        self.spell_search_var.trace_add('write', self.update_spell_list)
        
        self.spell_search_entry = ttk.Entry(search_spell_frame, textvariable=self.spell_search_var)
        self.spell_search_entry.pack(anchor=tk.NW, padx=10, pady=10)

        button_frame = ttk.Frame(self)
        button_frame.pack(anchor=tk.NW, padx=0, pady=10)

        self.add_spells_button = ttk.Button(button_frame, text="Add Spell", command=lambda: self.add_spells(self.get_spell_name()))
        self.add_spells_button.pack(side=tk.LEFT, padx=5)

        self.delete_spell_button = ttk.Button(button_frame, text="Delete Spell", command=lambda: self.delete_spell())
        self.delete_spell_button.pack(side=tk.LEFT, padx=5)

        self.edit_spell_button = ttk.Button(button_frame, text="Edit Spell", command=self.edit_spell)
        self.edit_spell_button.pack(side=tk.LEFT, padx=5)
        
        self.spell_details = ttk.Label(self, text="Select a spell to view details")
        self.spell_details.pack(anchor=tk.NW, padx=10, pady=10)

        self.spellbook = tk.Listbox(self, background="#F0ECE5", foreground="black", selectbackground="#161a30", selectforeground="white")
        self.spellbook.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.all_spell_data = hs.import_excel()
        
        self.spellbook.bind("<<ListboxSelect>>", self.show_spell_details)
        
        self.spell_levels = ["Cantrip", "1st Level", "2nd Level", "3rd Level", "4th Level", "5th Level", "6th Level", "7th Level", "8th Level", "9th Level"]
        self.spell_slots = [0, 4, 3, 3, 3, 2, 1, 1, 1, 1]
        self.notebook = ttk.Notebook(self, style="TNotebook")
        
        self.spell_level_listboxes = {}
        
        for i, level in enumerate(self.spell_levels):
            spell_level_frame = ttk.Frame(self.notebook, style="Frame1.TFrame")
            spell_level_frame.pack(fill=tk.BOTH, expand=True)
            
            if i == 0:
                spell_slots_label = ttk.Label(spell_level_frame, text="", style="TLabel1.TLabel")
            else:
                spell_slots_label = ttk.Label(spell_level_frame, text="Spell Slots: " + str(self.spell_slots[i]), style="TLabel1.TLabel")
            spell_slots_label.pack(side=tk.TOP, anchor=tk.W, padx=10, pady=10)
            
            cast_spell_button = ttk.Button(spell_level_frame, text="Cast Spell", command=lambda: self.cast_spell(), style="TButton1.TButton")
            cast_spell_button.pack(side=tk.TOP, anchor=tk.W, padx=10, pady=10)
            
            spell_level_listbox = tk.Listbox(spell_level_frame, background="#F0ECE5", foreground="black", selectbackground="#161a30", selectforeground="white")
            spell_level_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            self.spell_level_listboxes[level] = spell_level_listbox
            
            self.notebook.add(spell_level_frame, text=level)

        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.load_spell_list()
    
    def cast_spell(self):
        level = self.notebook.index(self.notebook.select())
        selected_tab = self.notebook.nametowidget(self.notebook.select())
        listbox = self.spell_level_listboxes[self.spell_levels[level]]
        selection = listbox.get(listbox.curselection())
        if selection:
            if self.spell_slots[level] > 0:
                self.spell_slots[level] -= 1
            for child in selected_tab.winfo_children():
                if isinstance(child, ttk.Label):
                    child.config(text="Spell Slots: " + str(self.spell_slots[level]))
                    
            self.save_spell_list()
        else:
            print("No spell selected")
        
    def update_spell_list(self, *args):
        spell_search = self.spell_search_var.get().lower()
        self.spellbook.delete(0, tk.END)
        for spell in self.spell_book_items:
            if spell_search.lower() in str(spell).lower():
                self.spellbook.insert(tk.END, spell)
    
    def get_spell_name(self):
        spellname = simpledialog.askstring("Spell Name", "Enter the name of the spell")
        return spellname
        
    def add_spells(self, spellname):
        for i in range(len(self.all_spell_data)):
            if str(spellname).lower() == str(self.all_spell_data[i]["Name"]).lower():
                self.spellbook.insert(tk.END, self.all_spell_data[i]["Name"])
                spell_level = self.all_spell_data[i]["Level"]
                self.notebook.select(spell_level)
                spell_level_frame = self.notebook.winfo_children()[int(spell_level)]
                spell_level_listbox = spell_level_frame.winfo_children()[2]
                spell_level_listbox.insert(tk.END, self.all_spell_data[i]["Name"]) 
                self.spell_book_items = self.spellbook.get(0, tk.END)
                
                self.save_spell_list()
            
                return
        
        self.spell_not_found()
        
        self.spell_book_items = self.spellbook.get(0, tk.END)
        self.save_spell_list()
    
    def spell_not_found(self):
        spell_not_found_window = tk.Toplevel(self)
        spell_not_found_window.title("Spell Not Found")
        spell_not_found_window.geometry("400x100")

        spell_not_found_label = ttk.Label(spell_not_found_window, text="Spell not found or is already in your spellbook. Please close this window and try again or create a custom spell.", wraplength=380)
        spell_not_found_label.pack(fill=tk.BOTH, expand=True, pady=10)

        create_custom_spell_button = ttk.Button(spell_not_found_window, text="Create Custom Spell", command=self.create_custom_spell)
        create_custom_spell_button.pack(fill=tk.BOTH, expand=True, pady=5)
    
    def edit_spell(self):
        spellname = self.spellbook.get(self.spellbook.curselection())
        for i in range(len(self.all_spell_data)):
            if spellname == self.all_spell_data[i]["Name"]:
                custom_spell_window = tk.Toplevel(self)
                custom_spell_window.title("Edit Spell")
                custom_spell_window.geometry("600x600")
                
                # Create a canvas to hold the widgets
                canvas = tk.Canvas(custom_spell_window)
                canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                
                # Create a scrollbar and associate it with the canvas
                scrollbar = ttk.Scrollbar(custom_spell_window, command=canvas.yview)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                canvas.configure(yscrollcommand=scrollbar.set)
                
                # Create a frame to contain the widgets
                frame = ttk.Frame(canvas)
                canvas.create_window((0, 0), window=frame, anchor=tk.NW)
                
                custom_spell_name_label = ttk.Label(frame, text="Spell Name")
                custom_spell_name_label.pack(pady=5)
                
                custom_spell_name_entry = ttk.Entry(frame)
                custom_spell_name_entry.insert(tk.END, self.all_spell_data[i]["Name"])
                custom_spell_name_entry.pack(pady=5)

                custom_spell_level_label = ttk.Label(frame, text="Spell Level")
                custom_spell_level_label.pack(pady=5)
                
                custom_spell_level_var = tk.StringVar()
                custom_spell_level_var.set(self.all_spell_data[i]["Level"])
                
                spell_levels = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
                spell_levels = list(map(str, spell_levels))
                custom_spell_level_menu = ttk.OptionMenu(frame, custom_spell_level_var, *spell_levels)
                custom_spell_level_menu.pack(pady=5)
                
                custom_spell_school_label = ttk.Label(frame, text="Spell School")
                custom_spell_school_label.pack(pady=5)
                
                custom_spell_school_var = tk.StringVar()
                custom_spell_school_var.set(self.all_spell_data[i]["School"])
                
                custom_spell_school_menu = ttk.OptionMenu(frame, custom_spell_school_var, "Abjuration", "Conjuration", "Divination", "Enchantment", "Evocation", "Illusion", "Necromancy", "Transmutation")
                custom_spell_school_menu.pack(pady=5)
                
                custom_spell_casting_time_label = ttk.Label(frame, text="Casting Time")
                custom_spell_casting_time_label.pack(pady=5)
                
                custom_spell_casting_time_entry = ttk.Entry(frame)
                custom_spell_casting_time_entry.insert(tk.END, self.all_spell_data[i]["Casting Time"])
                custom_spell_casting_time_entry.pack(pady=5)
                
                custom_spell_duration_label = ttk.Label(frame, text="Duration")
                custom_spell_duration_label.pack(pady=5)
                
                custom_spell_duration_entry = ttk.Entry(frame)
                custom_spell_duration_entry.insert(tk.END, self.all_spell_data[i]["Duration"])
                custom_spell_duration_entry.pack(pady=5)
                
                custom_spell_range_label = ttk.Label(frame, text="Range")
                custom_spell_range_label.pack(pady=5)
                
                custom_spell_range_entry = ttk.Entry(frame)
                custom_spell_range_entry.insert(tk.END, self.all_spell_data[i]["Range"])
                custom_spell_range_entry.pack(pady=5)
                
                custom_spell_area_label = ttk.Label(frame, text="Area")
                custom_spell_area_label.pack(pady=5)
                
                custom_spell_area_entry = ttk.Entry(frame)
                custom_spell_area_entry.insert(tk.END, self.all_spell_data[i]["Area"])
                custom_spell_area_entry.pack(pady=5)
                
                custom_spell_attack_label = ttk.Label(frame, text="Attack")
                custom_spell_attack_label.pack(pady=5)
                
                custom_spell_attack_entry = ttk.Entry(frame)
                custom_spell_attack_entry.insert(tk.END, self.all_spell_data[i]["Attack"])
                custom_spell_attack_entry.pack(pady=5)
                
                custom_spell_save_label = ttk.Label(frame, text="Save")
                custom_spell_save_label.pack(pady=5)
                
                custom_spell_save_entry = ttk.Entry(frame)
                custom_spell_save_entry.insert(tk.END, self.all_spell_data[i]["Save"])
                custom_spell_save_entry.pack(pady=5)
                
                custom_spell_damage_effect_label = ttk.Label(frame, text="Damage/Effect")
                custom_spell_damage_effect_label.pack(pady=5)
                
                custom_spell_damage_effect_entry = ttk.Entry(frame)
                custom_spell_damage_effect_entry.insert(tk.END, self.all_spell_data[i]["Damage/Effect"])
                custom_spell_damage_effect_entry.pack(pady=5)
                
                custom_spell_ritual_label = ttk.Label(frame, text="Ritual")
                custom_spell_ritual_label.pack(pady=5)
                
                custom_spell_ritual_var = tk.StringVar()
                custom_spell_ritual_var.set(self.all_spell_data[i]["Ritual"])
                
                custom_spell_ritual_menu = ttk.OptionMenu(frame, custom_spell_ritual_var, "Yes", "No")
                custom_spell_ritual_menu.pack(pady=5)
                
                custom_spell_concentration_label = ttk.Label(frame, text="Concentration")
                custom_spell_concentration_label.pack(pady=5)
                
                custom_spell_concentration_var = tk.StringVar()
                custom_spell_concentration_var.set(self.all_spell_data[i]["Concentration"])
                
                custom_spell_concentration_menu = ttk.OptionMenu(frame, custom_spell_concentration_var, "Yes", "No")
                custom_spell_concentration_menu.pack(pady=5)
                
                custom_spell_verbal_label = ttk.Label(frame, text="Verbal")
                custom_spell_verbal_label.pack(pady=5)
                
                custom_spell_verbal_var = tk.StringVar()
                custom_spell_verbal_var.set(self.all_spell_data[i]["Verbal"])
                
                custom_spell_verbal_menu = ttk.OptionMenu(frame, custom_spell_verbal_var, "Yes", "No")
                custom_spell_verbal_menu.pack(pady=5)
                
                custom_spell_somatic_label = ttk.Label(frame, text="Somatic")
                custom_spell_somatic_label.pack(pady=5)
                
                custom_spell_somatic_var = tk.StringVar()
                custom_spell_somatic_var.set(self.all_spell_data[i]["Somatic"])
                
                custom_spell_somatic_menu = ttk.OptionMenu(frame, custom_spell_somatic_var, "Yes", "No")
                custom_spell_somatic_menu.pack(pady=5)
                
                custom_spell_material_label = ttk.Label(frame, text="Material")
                custom_spell_material_label.pack(pady=5)
                
                custom_spell_material_var = tk.StringVar()
                custom_spell_material_var.set(self.all_spell_data[i]["Material"])
                
                custom_spell_material_menu = ttk.OptionMenu(frame, custom_spell_material_var, "Yes", "No")
                custom_spell_material_menu.pack(pady=5)
                
                custom_spell_material_star_label = ttk.Label(frame, text="Material*")
                custom_spell_material_star_label.pack(pady=5)
                
                custom_spell_material_star_entry = ttk.Entry(frame)
                custom_spell_material_star_entry.insert(tk.END, self.all_spell_data[i]["Material*"])
                custom_spell_material_star_entry.pack(pady=5)
                
                custom_spell_source_label = ttk.Label(frame, text="Source")
                custom_spell_source_label.pack(pady=5)
                
                custom_spell_source_entry = ttk.Entry(frame)
                custom_spell_source_entry.insert(tk.END, self.all_spell_data[i]["Source"])
                custom_spell_source_entry.pack(pady=5)
                
                custom_spell_details_label = ttk.Label(frame, text="Details")
                custom_spell_details_label.pack(pady=5)
                
                custom_spell_details_text = tk.Text(frame, wrap=tk.WORD)
                custom_spell_details_text.insert(tk.END, self.all_spell_data[i]["Details"])
                custom_spell_details_text.pack(pady=5)
                
                custom_spell_name_entry.insert(tk.END, " Edited Spell")
                
                custom_spell_save_button = ttk.Button(frame, text="Save Custom Spell", command=lambda: self.save_custom_spell(custom_spell_name_entry, custom_spell_level_var, custom_spell_school_var, custom_spell_casting_time_entry, custom_spell_duration_entry, custom_spell_range_entry, custom_spell_area_entry, custom_spell_attack_entry, custom_spell_save_entry, custom_spell_damage_effect_entry, custom_spell_ritual_var, custom_spell_concentration_var, custom_spell_verbal_var, custom_spell_somatic_var, custom_spell_material_var, custom_spell_material_star_entry, custom_spell_source_entry, custom_spell_details_text))
                custom_spell_save_button.pack(pady=5)
                
                # Update the canvas to match the size of the frame
                frame.update_idletasks()
                canvas.config(scrollregion=canvas.bbox("all"))

                for widget in frame.winfo_children():
                    widget.bind("<MouseWheel>", self.scroll_canvas)
                    
                break 
    
    def create_custom_spell(self):
        custom_spell_window = tk.Toplevel(self)
        custom_spell_window.title("Create Custom Spell")
        custom_spell_window.geometry("600x600")
        
        # Create a canvas to hold the widgets
        canvas = tk.Canvas(custom_spell_window)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create a scrollbar and associate it with the canvas
        scrollbar = ttk.Scrollbar(custom_spell_window, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create a frame to contain the widgets
        frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor=tk.NW)
        
        custom_spell_name_label = ttk.Label(frame, text="Spell Name")
        custom_spell_name_label.pack(pady=5)

        custom_spell_name_entry = ttk.Entry(frame)
        custom_spell_name_entry.pack(pady=5)

        custom_spell_level_label = ttk.Label(frame, text="Spell Level")
        custom_spell_level_label.pack(pady=5)

        custom_spell_level_var = tk.StringVar()
        custom_spell_level_var.set("0")

        spell_levels = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
        spell_levels = list(map(str, spell_levels))
        custom_spell_level_menu = ttk.OptionMenu(frame, custom_spell_level_var, *spell_levels)
        custom_spell_level_menu.pack(pady=5)

        custom_spell_school_label = ttk.Label(frame, text="Spell School")
        custom_spell_school_label.pack(pady=5)

        custom_spell_school_var = tk.StringVar()
        custom_spell_school_var.set("Abjuration")

        custom_spell_school_menu = ttk.OptionMenu(frame, custom_spell_school_var, "Abjuration", "Conjuration", "Divination", "Enchantment", "Evocation", "Illusion", "Necromancy", "Transmutation")
        custom_spell_school_menu.pack(pady=5)

        custom_spell_casting_time_label = ttk.Label(frame, text="Casting Time")
        custom_spell_casting_time_label.pack(pady=5)

        custom_spell_casting_time_entry = ttk.Entry(frame)
        custom_spell_casting_time_entry.pack(pady=5)

        custom_spell_duration_label = ttk.Label(frame, text="Duration")
        custom_spell_duration_label.pack(pady=5)

        custom_spell_duration_entry = ttk.Entry(frame)
        custom_spell_duration_entry.pack(pady=5)

        custom_spell_range_label = ttk.Label(frame, text="Range")
        custom_spell_range_label.pack(pady=5)

        custom_spell_range_entry = ttk.Entry(frame)
        custom_spell_range_entry.pack(pady=5)

        custom_spell_area_label = ttk.Label(frame, text="Area")
        custom_spell_area_label.pack(pady=5)

        custom_spell_area_entry = ttk.Entry(frame)
        custom_spell_area_entry.pack(pady=5)

        custom_spell_attack_label = ttk.Label(frame, text="Attack")
        custom_spell_attack_label.pack(pady=5)

        custom_spell_attack_entry = ttk.Entry(frame)
        custom_spell_attack_entry.pack(pady=5)

        custom_spell_save_label = ttk.Label(frame, text="Save")
        custom_spell_save_label.pack(pady=5)

        custom_spell_save_entry = ttk.Entry(frame)
        custom_spell_save_entry.pack(pady=5)

        custom_spell_damage_effect_label = ttk.Label(frame, text="Damage/Effect")
        custom_spell_damage_effect_label.pack(pady=5)

        custom_spell_damage_effect_entry = ttk.Entry(frame)
        custom_spell_damage_effect_entry.pack(pady=5)

        custom_spell_ritual_label = ttk.Label(frame, text="Ritual")
        custom_spell_ritual_label.pack(pady=5)

        custom_spell_ritual_var = tk.StringVar()
        custom_spell_ritual_var.set("No")

        custom_spell_ritual_menu = ttk.OptionMenu(frame, custom_spell_ritual_var, "Yes", "No")
        custom_spell_ritual_menu.pack(pady=5)

        custom_spell_concentration_label = ttk.Label(frame, text="Concentration")
        custom_spell_concentration_label.pack(pady=5)

        custom_spell_concentration_var = tk.StringVar()
        custom_spell_concentration_var.set("No")

        custom_spell_concentration_menu = ttk.OptionMenu(frame, custom_spell_concentration_var, "Yes", "No")
        custom_spell_concentration_menu.pack(pady=5)

        custom_spell_verbal_label = ttk.Label(frame, text="Verbal")
        custom_spell_verbal_label.pack(pady=5)

        custom_spell_verbal_var = tk.StringVar()
        custom_spell_verbal_var.set("No")

        custom_spell_verbal_menu = ttk.OptionMenu(frame, custom_spell_verbal_var, "Yes", "No")
        custom_spell_verbal_menu.pack(pady=5)

        custom_spell_somatic_label = ttk.Label(frame, text="Somatic")
        custom_spell_somatic_label.pack(pady=5)

        custom_spell_somatic_var = tk.StringVar()
        custom_spell_somatic_var.set("No")  

        custom_spell_somatic_menu = ttk.OptionMenu(frame, custom_spell_somatic_var, "Yes", "No")
        custom_spell_somatic_menu.pack(pady=5)

        custom_spell_material_label = ttk.Label(frame, text="Material")
        custom_spell_material_label.pack(pady=5)

        custom_spell_material_var = tk.StringVar()
        custom_spell_material_var.set("No")

        custom_spell_material_menu = ttk.OptionMenu(frame, custom_spell_material_var, "Yes", "No")
        custom_spell_material_menu.pack(pady=5)

        custom_spell_material_star_label = ttk.Label(frame, text="Material*")
        custom_spell_material_star_label.pack(pady=5)

        custom_spell_material_star_entry = ttk.Entry(frame)
        custom_spell_material_star_entry.pack(pady=5)

        custom_spell_source_label = ttk.Label(frame, text="Source")
        custom_spell_source_label.pack(pady=5)

        custom_spell_source_entry = ttk.Entry(frame)
        custom_spell_source_entry.pack(pady=5)

        custom_spell_details_label = ttk.Label(frame, text="Details")
        custom_spell_details_label.pack(pady=5)

        custom_spell_details_text = tk.Text(frame, wrap=tk.WORD)
        custom_spell_details_text.pack(pady=5)

        custom_spell_save_button = ttk.Button(frame, text="Save Custom Spell", command=lambda: self.save_custom_spell(custom_spell_name_entry, custom_spell_level_var, custom_spell_school_var, custom_spell_casting_time_entry, custom_spell_duration_entry, custom_spell_range_entry, custom_spell_area_entry, custom_spell_attack_entry, custom_spell_save_entry, custom_spell_damage_effect_entry, custom_spell_ritual_var, custom_spell_concentration_var, custom_spell_verbal_var, custom_spell_somatic_var, custom_spell_material_var, custom_spell_material_star_entry, custom_spell_source_entry, custom_spell_details_text))
        custom_spell_save_button.pack(pady=5)
        
        # Update the canvas to match the size of the frame
        frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))
        
        for widget in frame.winfo_children():
            widget.bind("<MouseWheel>", self.scroll_canvas)
            
    def scroll_canvas(self, event):
        # Get the delta value of the mouse wheel movement
        delta = event.delta

        # Scroll the canvas vertically based on the delta value
        self.canvas.yview_scroll(int(-1 * (delta / 120)), "units")
        
    def save_custom_spell(self, custom_spell_name_entry, custom_spell_level_var, custom_spell_school_var, custom_spell_casting_time_entry, custom_spell_duration_entry, custom_spell_range_entry, custom_spell_area_entry, custom_spell_attack_entry, custom_spell_save_entry, custom_spell_damage_effect_entry, custom_spell_ritual_var, custom_spell_concentration_var, custom_spell_verbal_var, custom_spell_somatic_var, custom_spell_material_var, custom_spell_material_star_entry, custom_spell_source_entry, custom_spell_details_text):    
        custom_spell = {
            "Unnamed: 0": "Custom Spell",
            "Name": custom_spell_name_entry.get(),
            "Level": custom_spell_level_var.get(),
            "School": custom_spell_school_var.get(),
            "Casting Time": custom_spell_casting_time_entry.get(),
            "Duration": custom_spell_duration_entry.get(),
            "Range": custom_spell_range_entry.get(),
            "Area": custom_spell_area_entry.get(),
            "Attack": custom_spell_attack_entry.get(),
            "Save": custom_spell_save_entry.get(),
            "Damage/Effect": custom_spell_damage_effect_entry.get(),
            "Ritual": custom_spell_ritual_var.get(),
            "Concentration": custom_spell_concentration_var.get(),
            "Verbal": custom_spell_verbal_var.get(),
            "Somatic": custom_spell_somatic_var.get(),
            "Material": custom_spell_material_var.get(),
            "Material*": custom_spell_material_star_entry.get(),
            "Source": custom_spell_source_entry.get(),
            "Details": custom_spell_details_text.get("1.0", tk.END),
            "Unnamed: 19": "Custom Spell"
        }
        
        hs.add_to_excel(custom_spell)
        self.all_spell_data = hs.import_excel()
        self.add_spells(custom_spell_name_entry.get())
      
    def delete_spell(self):
        spellname = self.spellbook.get(self.spellbook.curselection())
        self.spellbook.delete(self.spellbook.curselection())
        for i in range(len(self.all_spell_data)):
            if spellname == self.all_spell_data[i]["Name"]:
                spell_level = self.all_spell_data[i]["Level"]
                self.notebook.select(spell_level)
                spell_level_frame = self.notebook.winfo_children()[int(spell_level)]
                spell_level_listbox = spell_level_frame.winfo_children()[2]
                spell_level_listbox.delete(spell_level_listbox.get(0, tk.END).index(spellname))
                break
            
        self.spell_book_items = self.spellbook.get(0, tk.END)
        self.save_spell_list()

    def show_spell_details(self, event):
        spellname = self.spellbook.get(self.spellbook.curselection())

        for i in range(len(self.all_spell_data)):
            if spellname == self.all_spell_data[i]["Name"]:
                # new window with spell details
                spell_info_window = tk.Toplevel(self)
                spell_info_window.title(self.all_spell_data[i]["Name"])
                spell_info_window.geometry("500x400")
                print(self.all_spell_data[i]["Name"])

                # Create a scrolled text widget
                spell_info_text = scrolledtext.ScrolledText(spell_info_window, wrap=tk.WORD)
                spell_info_text.pack(fill=tk.BOTH, expand=True)

                # Set the spell details text
                spell_info_text.insert(tk.END, "Level: " + str(self.all_spell_data[i]["Level"]) + "\n")
                spell_info_text.insert(tk.END, "School: " + str(self.all_spell_data[i]["School"]) + "\n")
                spell_info_text.insert(tk.END, "Casting Time: " + str(self.all_spell_data[i]["Casting Time"]) + "\n")
                spell_info_text.insert(tk.END, "Duration: " + str(self.all_spell_data[i]["Duration"]) + "\n")
                spell_info_text.insert(tk.END, "Range: " + str(self.all_spell_data[i]["Range"]) + "\n")
                spell_info_text.insert(tk.END, "Area: " + str(self.all_spell_data[i]["Area"]) + "\n")
                spell_info_text.insert(tk.END, "Attack: " + str(self.all_spell_data[i]["Attack"]) + "\n")
                spell_info_text.insert(tk.END, "Save: " + str(self.all_spell_data[i]["Save"]) + "\n")
                spell_info_text.insert(tk.END, "Damage/Effect: " + str(self.all_spell_data[i]["Damage/Effect"]) + "\n")
                spell_info_text.insert(tk.END, "Ritual: " + str(self.all_spell_data[i]["Ritual"]) + "\n")
                spell_info_text.insert(tk.END, "Concentration: " + str(self.all_spell_data[i]["Concentration"]) + "\n")
                spell_info_text.insert(tk.END, "Verbal: " + str(self.all_spell_data[i]["Verbal"]) + "\n")
                spell_info_text.insert(tk.END, "Somatic: " + str(self.all_spell_data[i]["Somatic"]) + "\n")
                spell_info_text.insert(tk.END, "Material: " + str(self.all_spell_data[i]["Material"]) + "\n")
                spell_info_text.insert(tk.END, "Material*: " + str(self.all_spell_data[i]["Material*"]) + "\n")
                spell_info_text.insert(tk.END, "Source: " + str(self.all_spell_data[i]["Source"]) + "\n")
                spell_info_text.insert(tk.END, "Details: " + "\n")
                spell_info_text.insert(tk.END, str(self.all_spell_data[i]["Details"]))

                break
    
    def save_spell_list(self):
        spell_list = self.spellbook.get(0, tk.END)
        spell_slots = self.spell_slots
        
        data = {
            "Spells": spell_list,
            "Spell Slots": spell_slots,
        }
        
        with open("spell_list.json", "w") as f:
            json.dump(data, f)
        print("Spell list saved to spell_list.json")
        
    def load_spell_list(self):
        try:
            with open("spell_list.json", "r") as f:
                data = json.load(f)
                
                self.spell_slots = data["Spell Slots"]
                for i in range(len(self.spell_slots)):
                    if i != 0:
                        self.notebook.winfo_children()[i].winfo_children()[0].config(text="Spell Slots: " + str(self.spell_slots[i]), style="TLabel1.TLabel")
                
                for spell in data["Spells"]:
                    self.add_spells(spell)
                print("Spell list loaded from spell_list.json")
        except FileNotFoundError:
            print("Spell list file not found. Starting with an empty spell list.")

if __name__ == "__main__":
    app = SpellbookApp()
    app.mainloop()