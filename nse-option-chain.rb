require 'net/http'
require 'json'
require 'csv'
require 'terminal-table'
require 'write_xlsx'

def fetch_nifty_option_chain
  url = URI("https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY")
  headers = {
    "User-Agent" => "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Accept-Language" => "en-US,en;q=0.9",
    "Accept" => "application/json"
  }

  http = Net::HTTP.new(url.host, url.port)
  http.use_ssl = true
  http.open_timeout = 10 # Timeout for opening the connection
  http.read_timeout = 20 # Timeout for reading the response

  request = Net::HTTP::Get.new(url, headers)

  response = http.request(request)

  if response.code == "200"
    JSON.parse(response.body)
  else
    puts "Error fetching data: #{response.code}"
    nil
  end
end

# Function to filter data for a specific expiry date
def filter_by_expiry(data, expiry_date)
  return [] unless data && data["records"] && data["records"]["data"]

  data["records"]["data"].select do |record|
    record["expiryDate"] == expiry_date
  end
end

def fetch_nifty_option_chain_with_retries(max_retries = 3)
  retries = 0
  puts "================== #{Time.now} ===================="
  begin
    fetch_nifty_option_chain
  rescue Net::ReadTimeout => e
    retries += 1
    if retries <= max_retries
      puts "Timeout occurred. Retrying... (Attempt #{retries})"
      sleep(2)
      retry
    else
      puts "Failed after #{max_retries} retries: #{e.message}"
      nil
    end
  end
end

# Function to save or update CSV file
def save_or_update_csv(file_path, filtered_data)
  # Ensure file exists, otherwise create with headers
  unless File.exist?(file_path)
    CSV.open(file_path, "w") do |csv|
      csv << ["Time", "Strike Price", "CE OI", "CE Change OI", "CE Volume", "PE OI", "PE Change OI", "PE Volume"]
    end
  end

  # Load existing data for updates
  existing_data = {}
  CSV.foreach(file_path, headers: true) do |row|
    existing_data[row["Strike Price"]] = row.to_h
  end

  # Update or append new data
  CSV.open(file_path, "wb") do |csv|
    # Write headers
    csv << ["Time", "CE OI", "CE Change OI", "CE Volume", "Strike Price", "PE OI", "PE Change OI", "PE Volume"]

    # Merge existing data with new data
    filtered_data.each do |record|
      strike_price = record["strikePrice"].to_s
      ce_data = record["CE"] || {}
      pe_data = record["PE"] || {}

      row = {
        "Time" => Time.now.strftime("%H:%M:%S"),
        "CE OI" => ce_data["openInterest"],
        "CE Change OI" => ce_data["changeinOpenInterest"],
        "CE Volume" => ce_data["totalTradedVolume"],
        "Strike Price" => strike_price,
        "PE OI" => pe_data["openInterest"],
        "PE Change OI" => pe_data["changeinOpenInterest"],
        "PE Volume" => pe_data["totalTradedVolume"]
      }

      # Update existing data or add new
      existing_data[strike_price] = row
    end

    # Write all updated data back to CSV
    existing_data.each_value do |row|
      csv << row.values
    end
  end
end

# Function to save Excel file with highlights
def save_excel(file_path, filtered_data)
  workbook = WriteXLSX.new(file_path)
  worksheet = workbook.add_worksheet

  # Add headers
  headers = ["CE OI", "CE Change OI", "CE Volume", "Strike Price", "PE OI", "PE Change OI", "PE Volume"]
  worksheet.write_row(0, 0, headers)

  # Identify max OI for highlighting
  ce_ois = filtered_data.map { |record| (record["CE"] || {})["openInterest"] }
  pe_ois = filtered_data.map { |record| (record["PE"] || {})["openInterest"] }
  max_ce_oi = ce_ois.compact.max
  max_pe_oi = pe_ois.compact.max

  # Add data with highlights
  filtered_data.each_with_index do |record, row_index|
    ce_data = record["CE"] || {}
    pe_data = record["PE"] || {}
    row_data = [
      ce_data["openInterest"],
      ce_data["changeinOpenInterest"],
      ce_data["totalTradedVolume"],
      record["strikePrice"],
      pe_data["openInterest"],
      pe_data["changeinOpenInterest"],
      pe_data["totalTradedVolume"]
    ]

    row_data.each_with_index do |value, col_index|
      format = workbook.add_format
      if col_index == 1 && value == max_ce_oi
        format.set_bg_color('red')
      elsif col_index == 4 && value == max_pe_oi
        format.set_bg_color('green')
      end
      worksheet.write(row_index + 1, col_index, value, format)
    end
  end

  workbook.close
  puts "Data saved to #{file_path}"
end

require 'terminal-table'

# Function to print data in terminal with top 3 support and resistance highlighted
def print_table(filtered_data)
  # Extract rows
  rows = filtered_data.map do |record|
    ce_data = record["CE"] || {}
    pe_data = record["PE"] || {}
    [
      ce_data["openInterest"],
      ce_data["changeinOpenInterest"],
      ce_data["totalTradedVolume"],
      record["strikePrice"], 
      pe_data["openInterest"],
      pe_data["changeinOpenInterest"],
      pe_data["totalTradedVolume"]
    ]
  end

  # Identify top 3 support (PE OI) and resistance (CE OI) levels
  top_resistance = rows.sort_by { |row| -row[0].to_i }.first(3) # Sort by CE OI descending
  top_support = rows.sort_by { |row| -row[4].to_i }.first(3)    # Sort by PE OI descending

  # Define color codes
  colors = {
    resistance: ["\e[31m", "\e[33m", "\e[35m"], # Red, Yellow, Magenta for resistance
    support: ["\e[32m", "\e[34m", "\e[36m"]    # Green, Blue, Cyan for support
  }

  # Highlight rows for top 3 resistance and support
  rows.each do |row|
    if top_resistance.include?(row)
      index = top_resistance.index(row)
      row[0] = "#{colors[:resistance][index]}#{row[0]}\e[0m"
    end

    if top_support.include?(row)
      index = top_support.index(row)
      row[4] = "#{colors[:support][index]}#{row[4]}\e[0m"
    end
  end

  # Create and print the table with strike price centered
  table = Terminal::Table.new(
    title: "Nifty Option Chain (Filtered)",
    headings: ["CE OI", "CE Change OI", "CE Volume","Strike Price", "PE OI", "PE Change OI", "PE Volume"],
    rows: rows
  )

  # Center the strike price column (index 0)
  table.align_column(0, :center)

  puts table
end


# Identify ATM Strike Price and capture before/after range with expiry filter
def get_atm_and_nearby_strikes(data, expiry_date, range = 10)
  return [] unless data && data["records"] && data["records"]["data"]
  spot_price = data.dig("records", "underlyingValue") # Fetch NIFTY Spot Price
  all_strike_prices = data["records"]["data"].map { |record| record["strikePrice"] }.uniq.sort

  # Find ATM Strike Price
  atm_strike_price = all_strike_prices.min_by { |strike| (strike - spot_price).abs }
  puts "Spot Price: #{spot_price}, ATM Strike Price: #{atm_strike_price}"

  # Get strike prices in range of ATM
  start_index = [all_strike_prices.index(atm_strike_price) - range, 0].max
  end_index = [all_strike_prices.index(atm_strike_price) + range, all_strike_prices.size - 1].min
  selected_strike_prices = all_strike_prices[start_index..end_index]

  # Filter option chain data by expiry date and strike prices
  data["records"]["data"].select do |record|
    record["expiryDate"] == expiry_date && selected_strike_prices.include?(record["strikePrice"])
  end
end

def decide_option_strategy(filtered_data)
  best_option = nil
  best_score = -Float::INFINITY  # Initialize with negative infinity to track the best score
  
  filtered_data.each do |option|
    strike_price = option["strikePrice"]
    ce_data = option["CE"]
    pe_data = option["PE"]
    
    # Skip options if no data exists for CE or PE
    next if ce_data.nil? || pe_data.nil?
    
    # Extract necessary values for CE and PE
    ce_ltp = ce_data["lastPrice"]
    pe_ltp = pe_data["lastPrice"]
    ce_oi = ce_data["openInterest"]
    pe_oi = pe_data["openInterest"]
    ce_iv = ce_data["impliedVolatility"]
    pe_iv = pe_data["impliedVolatility"]
    
    # Skip invalid data
    next if ce_ltp.nil? || pe_ltp.nil? || ce_oi.nil? || pe_oi.nil? || ce_iv.nil? || pe_iv.nil?

    # Calculate option "score" based on OI and IV
    ce_score = ce_oi * ce_iv  # This combines OI and IV for CE
    pe_score = pe_oi * pe_iv  # This combines OI and IV for PE
    
    # Assuming the current price is the underlying value for comparison
    current_price = option["CE"]["underlyingValue"] # Using the underlying value (e.g., NIFTY index)

    # Determine market direction (bullish/bearish)
    if current_price > strike_price # Market is bullish, focus on CE
      # We compare the "score" of the CE options
      if ce_score > best_score
        best_score = ce_score
        
        # Dynamically adjust the ranges based on volatility and momentum
        buy_range_min = (ce_ltp * 0.995).round(2)  # 0.5% below LTP
        buy_range_max = (ce_ltp * 1.005).round(2)  # 0.5% above LTP
        
        # Adjust stop loss more dynamically based on volatility
        stop_loss_min = (ce_ltp * 0.97).round(2)  # 3% below LTP
        stop_loss_max = (ce_ltp * 0.98).round(2)  # 2% below LTP

        # Increase target range to maximize profit potential
        target_min = (ce_ltp * 1.05).round(2)  # 5% above LTP
        target_max = (ce_ltp * 1.10).round(2)  # 10% above LTP
        
        # Set buffer values dynamically based on volatility
        if ce_iv > 25  # High volatility, higher target
          target_min = (ce_ltp * 1.10).round(2)  # 10% above LTP
          target_max = (ce_ltp * 1.15).round(2)  # 15% above LTP
        elsif ce_iv < 15  # Low volatility, smaller target
          target_min = (ce_ltp * 1.03).round(2)  # 3% above LTP
          target_max = (ce_ltp * 1.05).round(2)  # 5% above LTP
        end
        
        best_option = {
          type: "CE",
          strike_price: strike_price,
          buy_range_min: buy_range_min,
          buy_range_max: buy_range_max,
          stop_loss: "#{stop_loss_min}-#{stop_loss_max}",
          target: "#{target_min}-#{target_max}",
          score: ce_score
        }
      end
    elsif current_price < strike_price # Market is bearish, focus on PE
      # We compare the "score" of the PE options
      if pe_score > best_score
        best_score = pe_score
        
        # Dynamically adjust the ranges based on volatility and momentum
        buy_range_min = (pe_ltp * 0.995).round(2)  # 0.5% below LTP
        buy_range_max = (pe_ltp * 1.005).round(2)  # 0.5% above LTP
        
        # Adjust stop loss more dynamically based on volatility
        stop_loss_min = (pe_ltp * 1.02).round(2)  # 2% above LTP
        stop_loss_max = (pe_ltp * 1.03).round(2)  # 3% above LTP

        # Increase target range to maximize profit potential
        target_min = (pe_ltp * 0.95).round(2)  # 5% below LTP
        target_max = (pe_ltp * 0.90).round(2)  # 10% below LTP
        
        # Set buffer values dynamically based on volatility
        if pe_iv > 25  # High volatility, higher target
          target_min = (pe_ltp * 0.90).round(2)  # 10% below LTP
          target_max = (pe_ltp * 0.85).round(2)  # 15% below LTP
        elsif pe_iv < 15  # Low volatility, smaller target
          target_min = (pe_ltp * 0.98).round(2)  # 2% below LTP
          target_max = (pe_ltp * 0.95).round(2)  # 5% below LTP
        end
        
        best_option = {
          type: "PE",
          strike_price: strike_price,
          buy_range_min: buy_range_min,
          buy_range_max: buy_range_max,
          stop_loss: "#{stop_loss_min}-#{stop_loss_max}",
          target: "#{target_min}-#{target_max}",
          score: pe_score
        }
      end
    end
  end

  # Output the most profitable option based on OI and IV
  if best_option
    puts "Most Profitable Option:"
    puts "#{best_option[:type]} - Strike Price: #{best_option[:strike_price]}"
    puts "BUY Range: #{best_option[:buy_range_min]}-#{best_option[:buy_range_max]}"
    puts "Stop Loss Range: #{best_option[:stop_loss]}"
    puts "Target Range: #{best_option[:target]}"
    puts "OI: #{best_option[:score]}"
  else
    puts "No profitable option found based on OI and IV."
  end
end


# Main Execution
begin
  file_path = "/Users/gowri/Downloads/nifty_option_chain.csv"
  file_path_excel = "/Users/gowri/Downloads/nifty_option_chain.xlsx"
  expiry_date = nil

  loop do
    # Fetch option chain data
    data = fetch_nifty_option_chain_with_retries

    if data
      # Set expiry date if not already set
      expiry_date ||= data.dig("records", "expiryDates", 0) # First available expiry date
      puts "Using expiry date: #{expiry_date}"

      # Filter data for specific expiry
      filtered_data = get_atm_and_nearby_strikes(data, expiry_date)

      # Save or update data in CSV
      # save_or_update_csv(file_path, filtered_data)
      # save_excel(file_path_excel, filtered_data)
      print_table(filtered_data)
      # puts "Data updated in #{file_path}"
      decide_option_strategy(filtered_data)
    end

    # Wait for 3 minutes
    sleep(120)
  end
rescue StandardError => e
  puts "An error occurred: #{e.message}"
  puts e.backtrace
end