require 'watir'
require 'faker'
require 'spreadsheet'

browser = Watir::Browser.new
Spreadsheet.client_encoding = 'UTF-8'
information = Spreadsheet::Workbook.new
sheet1 = information.create_worksheet :name => 'JVVNL'
sheet1.row(0).push  'Name', 'K Number', 'Binder No', 'Account No', 'Amount'
10.times do |n|
  @url = browser.goto('https://www.billdesk.com/pgidsk/pgmerc/jvvnljp/JVVNLJPDetails.jsp?billerid=RVVNLJP')
  k_num = (210742023980+n).to_s
  form = browser.form(name: 'form1')
  form.radio(name: 'service',:value => 'BILL').set
  form.text_field(name: 'txtCustomerID').set(k_num)
  form.text_field(name: 'txtEmail').set(Faker::Internet.email)
  form.button(class_name: 'subtn').click
  browser.screenshot.save ("preview_browser#{n}.png")
  if browser.table.id == 'tb_confirm'
    amount_payable = browser.tr(text: /Amount Payable/).td(:index => 1).text
    k_number = browser.tr(text: /K Number/).td(:index => 1).text
    binder_number = browser.tr(text: /Binder Number/).td(:index => 1).text
    account_number = browser.tr(text: /Account Number/).td(:index => 1).text
    customer_name = browser.tr(text: /Customer Name/).td(:index => 1).text
    sheet1.row(sheet1.last_row_index+1).push customer_name, k_number, binder_number, account_number, amount_payable if amount_payable.to_i > 10000
    information.write 'JVVNL.xls'
  elsif browser.table.id == 'tb_detail'
    @url
  end
end
browser.close()
