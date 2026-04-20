import tkinter as tk
from tkinter import ttk
import urllib.request
import json
import time
import threading

class GEMixin:
    def _calculate_ge_price(self, action_type, high_price, low_price):
        """Calculates a final price based on the strategy for either 'buy' or 'sell'."""
        if action_type == 'buy':
            strategy = self.ge_interface_buy_price_strategy.get()
            custom_price = int(self.ge_interface_buy_custom_price.get())
            margin = int(self.ge_interface_buy_price_margin.get())
            
            if strategy == 'Insta-Buy': return high_price
            if strategy == '+5%': return int(high_price * 1.05)
            if strategy == '-5%': return int(high_price * 0.95)
            if strategy == 'Custom Price': return custom_price
            if strategy == 'Flip-Buy (use Insta-Sell)': return low_price
            if strategy == 'Flip-Buy (Insta-Sell + Margin)': return low_price + margin
            return high_price # Default
        else: # sell
            strategy = self.ge_interface_sell_price_strategy.get()
            custom_price = int(self.ge_interface_sell_custom_price.get())
            margin = int(self.ge_interface_sell_price_margin.get())
            
            if strategy == 'Insta-Sell': return low_price
            if strategy == '+5%': return int(low_price * 1.05)
            if strategy == '-5%': return int(low_price * 0.95)
            if strategy == 'Custom Price': return custom_price
            if strategy == 'Flip-Sell (use Insta-Buy)': return high_price
            if strategy == 'Flip-Sell (Insta-Buy - Margin)': return high_price - margin
            return low_price # Default

    def _toggle_ge_buy_options(self, *args):
        strategy = self.ge_interface_buy_price_strategy.get()
        self.ge_buy_custom_price_entry.pack_forget()
        self.ge_buy_margin_entry.pack_forget()
        if strategy == 'Custom Price':
            self.ge_buy_custom_price_entry.pack(padx=5, pady=(0,5), fill=tk.X, expand=True)
        elif 'Margin' in strategy:
            self.ge_buy_margin_entry.pack(padx=5, pady=(0,5), fill=tk.X, expand=True)

    def _toggle_ge_sell_options(self, *args):
        strategy = self.ge_interface_sell_price_strategy.get()
        self.ge_sell_custom_price_entry.pack_forget()
        self.ge_sell_margin_entry.pack_forget()
        if strategy == 'Custom Price':
            self.ge_sell_custom_price_entry.pack(padx=5, pady=(0,5), fill=tk.X, expand=True)
        elif 'Margin' in strategy:
            self.ge_sell_margin_entry.pack(padx=5, pady=(0,5), fill=tk.X, expand=True)

    def _process_ge_price_data_on_main_thread(self, price_data):
        """Processes the fetched GE data and updates the UI. Runs in the main thread."""
        self.ge_interface_last_data = price_data
        try:
            if not isinstance(price_data, dict):
                raise ValueError("API returned no valid price data.")

            high_price = int(price_data.get('high'))
            low_price = int(price_data.get('low'))
            quantity = int(self.ge_interface_item_quantity.get())
            item_name = self.ge_interface_item_name.get()

            # Calculate final prices based on independent strategies
            final_buy_price = self._calculate_ge_price('buy', high_price, low_price)
            final_sell_price = self._calculate_ge_price('sell', high_price, low_price)

            # Update UI display
            self.ge_interface_display_buy_price.set(f"{final_buy_price:,}")
            self.ge_interface_display_buy_total.set(f"{final_buy_price * quantity:,}")
            self.ge_interface_display_sell_price.set(f"{final_sell_price:,}")
            self.ge_interface_display_sell_total.set(f"{final_sell_price * quantity:,}")
            
            self.log(f"GE Interface: Updated prices for {item_name}.")

        except (ValueError, TypeError, KeyError) as e:
            self.ge_interface_display_buy_price.set("Error")
            self.ge_interface_display_buy_total.set("Error")
            self.ge_interface_display_sell_price.set("Error")
            self.ge_interface_display_sell_total.set("Error")
            self.log(f"GE Interface: Failed to process data for '{self.ge_interface_item_name.get()}'. Reason: {e}", "red")

    def _toggle_ge_auto_update(self, *args):
        if self.ge_auto_update_enabled.get():
            self._start_ge_auto_updater()
        else:
            self._stop_ge_auto_updater()

    def _start_ge_auto_updater(self):
        self._stop_ge_auto_updater() # Ensure no duplicates
        self.log("GE auto-updater started.", "green")
        self._run_ge_auto_update_cycle()

    def _stop_ge_auto_updater(self):
        if self.ge_auto_update_after_id:
            self.root.after_cancel(self.ge_auto_update_after_id)
            self.ge_auto_update_after_id = None
            self.log("GE auto-updater stopped.")

    def _run_ge_auto_update_cycle(self):
        if not self.ge_auto_update_enabled.get():
            self._stop_ge_auto_updater()
            return

        self.update_ge_interface_price()
        try:
            interval_ms = int(self.ge_auto_update_interval.get()) * 1000
            if interval_ms > 0:
                 self.ge_auto_update_after_id = self.root.after(interval_ms, self._run_ge_auto_update_cycle)
            else:
                 self.log("Auto-update interval must be greater than 0.", "orange")
                 self.ge_auto_update_enabled.set(False)
        except ValueError:
            self.log("Invalid auto-update interval. Please enter a number.", "orange")
            self.ge_auto_update_enabled.set(False)

    def _toggle_ge_interface_price_options(self, *args):
        strategy = self.ge_interface_price_strategy.get()
        self.ge_custom_price_entry.pack_forget()
        self.ge_margin_entry.pack_forget()
        if strategy == 'Custom Price':
            self.ge_custom_price_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        elif 'Margin' in strategy:
            self.ge_margin_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
    def get_item_mapping(self):
        if self.item_mapping_cache is not None: return self.item_mapping_cache
        url = "https://prices.runescape.wiki/api/v1/osrs/mapping"
        try:
            req = urllib.request.Request(url, headers=self.api_headers)
            with urllib.request.urlopen(req) as response:
                if response.status == 200:
                    data = json.loads(response.read().decode())
                    self.item_mapping_cache = {'by_name': {item['name'].lower(): item for item in data}, 'by_id': {item['id']: item for item in data}}
                    self.log("Successfully downloaded and cached OSRS item map.")
                    return self.item_mapping_cache
                else: self.log(f"API Error: Failed to get item map (Status: {response.status})", "red"); return None
        except Exception as e: self.log(f"API Request Error: {e}", "red"); return None

    def get_item_price(self, item_name):
        item_map_data = self.get_item_mapping()
        if not item_map_data:
            return None
        item_map = item_map_data.get('by_name')
        if not item_map:
            return None

        item_data = item_map.get(item_name.lower())
        if not item_data:
            self.log(f"Item '{item_name}' not found in the mapping.", "orange")
            return None
        
        item_id = item_data['id']
        if item_id in self.item_price_cache:
            cached_data, timestamp = self.item_price_cache[item_id]
            if (time.time() - timestamp) < 60:
                return cached_data
            
        url = f"https://prices.runescape.wiki/api/v1/osrs/latest?id={item_id}"
        try:
            req = urllib.request.Request(url, headers=self.api_headers)
            with urllib.request.urlopen(req) as response:
                if response.status == 200:
                    response_json = json.loads(response.read().decode())
                    if not response_json or 'data' not in response_json or response_json['data'] is None:
                        self.log(f"API Error: Malformed data response for {item_name}", "red")
                        return None

                    data = response_json['data']
                    if str(item_id) in data:
                        item_price_info = data[str(item_id)]
                        self.item_price_cache[item_id] = (item_price_info, time.time())
                        return item_price_info
                    else:
                        self.log(f"API Warning: Price data not available for {item_name} (ID: {item_id})", "orange")
                        return None
                else:
                    self.log(f"API Error: Failed to get price for {item_name} (Status: {response.status})", "red")
                    return None
        except Exception as e:
            self.log(f"API Request Error for {item_name}: {e}", "red")
            return None

    def get_all_latest_prices(self):
        if self.all_item_prices_cache:
            cached_data, timestamp = self.all_item_prices_cache
            if (time.time() - timestamp) < 300: return cached_data
        
        url = "https://prices.runescape.wiki/api/v1/osrs/latest"
        try:
            req = urllib.request.Request(url, headers=self.api_headers)
            with urllib.request.urlopen(req) as response:
                if response.status == 200:
                    data = json.loads(response.read().decode())['data']
                    self.all_item_prices_cache = (data, time.time()); self.log("Successfully downloaded latest prices for all items.")
                    return data
                else: self.log(f"API Error: Failed to get all prices (Status: {response.status})", "red"); return None
        except Exception as e: self.log(f"API Request Error (all prices): {e}", "red"); return None

    def get_all_hourly_volumes(self):
        if self.hourly_volume_cache:
            cached_data, timestamp = self.hourly_volume_cache
            if (time.time() - timestamp) < 300: return cached_data

        url = "https://prices.runescape.wiki/api/v1/osrs/1h"
        try:
            req = urllib.request.Request(url, headers=self.api_headers)
            with urllib.request.urlopen(req) as response:
                if response.status == 200:
                    data = json.loads(response.read().decode())['data']
                    self.hourly_volume_cache = (data, time.time())
                    self.log("Successfully downloaded 1-hour volumes for all items.")
                    return data
                else: self.log(f"API Error: Failed to get all volumes (Status: {response.status})", "red"); return None
        except Exception as e: self.log(f"API Request Error (all volumes): {e}", "red"); return None
        
    def update_ge_interface_price(self):
        item_name = self.ge_interface_item_name.get()
        if not item_name:
            self.log("GE Interface: Please enter an item name.", "orange")
            return
        
        # Set a loading state in the UI immediately
        self.ge_interface_display_buy_price.set("Fetching...")
        self.ge_interface_display_buy_total.set("Fetching...")
        self.ge_interface_display_sell_price.set("Fetching...")
        self.ge_interface_display_sell_total.set("Fetching...")

        # Run the network request in a separate thread to avoid freezing the GUI
        fetch_thread = threading.Thread(target=self._fetch_ge_price_in_thread, daemon=True)
        fetch_thread.start()

    def _fetch_ge_price_in_thread(self):
        """Wrapper to run get_item_price in a thread and schedule UI update."""
        item_name = self.ge_interface_item_name.get()
        price_data = self.get_item_price(item_name)
        # Safely schedule the UI update to run on the main thread
        self.root.after(0, self._process_ge_price_data_on_main_thread, price_data)

