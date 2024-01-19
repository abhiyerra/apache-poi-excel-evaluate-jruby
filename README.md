# apache-poi-excel-evaluate

An example to evaluate excel files using JRuby and Java

wget https://repo1.maven.org/maven2/org/jruby/jruby-dist/9.4.5.0/jruby-dist-9.4.5.0-bin.tar.gz
tar zxvf jruby-dist-9.4.5.0-bin.tar.gz
mv jruby-9.4.5.0 jruby

wget https://dlcdn.apache.org/poi/release/src/poi-src-5.2.5-20231118.tgz
tar zxvf poi-src-5.2.5-20231118.tgz
mv poi-src-5.2.5-20231118 poi

wget https://archive.apache.org/dist/poi/release/bin/poi-bin-5.2.3-20220909.zip
unzip -a ./poi-bin-5.2.3-20220909.zip

./jruby/bin/jruby ./run-formula.rb
