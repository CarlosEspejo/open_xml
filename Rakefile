require "rake/testtask"

Rake::TestTask.new do |t|
  t.libs << "lib"
  t.libs << "spec"
  t.pattern = "spec/**/*_spec.rb"
end

task :default => :test
