## Convert MySQL Database to Microsoft Excel

## Usage

first step,

```
git clone https://github.com/kzf3hase/export_db_difinition_document.git
```

second step, copy database.rb.sample and rename database.rb

```
cp config/database.rb.sample config/database.rb
```

third step, set up database information

```
# config/database.rb

$database = {
  :host => "YOUR DATABASE HOST",
  :username => "YOUR DATABASE USER NAME",
  :password => "YOUR PASSWORD",
  :database => "YOUR DATABASE"
}
```

run! 


```
ruby export.rb {FILE_NAME}
```

Specify FILE_NAME is without extension


if outpu except curent directory.

```
ruby export.rb /tmp/{FILE_NAME}
```
