use strict;
use warnings;
use File::Find;
use File::Spec;

# Get command line arguments: base directory, directory output file, and file list output file
my ($base_dir, $dir_output, $file_output) = @ARGV;

# Validate input arguments
die "Usage: $0 <base_directory> <dir_output_file> <file_output_file>\n" 
  unless defined $base_dir && defined $dir_output && defined $file_output;
die "Base directory '$base_dir' does not exist or is not a directory\n" 
  unless -d $base_dir;

my %dirs;     # Hash to store unique directory paths
my @files;    # Array to store all .c file paths

# Traverse the base directory to find all .c files
find({
    wanted => sub {
        if (-f $_ && $_ =~ /\.c$/i) {  # Check .c files (case-insensitive)
            # 1. Process directories
            my $abs_dir = $File::Find::dir;
            my $rel_dir = File::Spec->abs2rel($abs_dir, $base_dir);
            $rel_dir =~ s[/][\\]g;  # Convert to Windows-style path
            $dirs{$rel_dir} = 1;

            # 2. Process files
            my $abs_file = $File::Find::name;
            my $rel_file = File::Spec->abs2rel($abs_file, $base_dir);
            $rel_file =~ s[/][\\]g;  # Convert to Windows-style path
            push @files, $rel_file;
        }
    },
    no_chdir => 1  # Maintain accurate path tracking
}, $base_dir);

# Write directory paths
open(my $dir_fh, '>', $dir_output) 
  or die "Cannot open directory output file '$dir_output': $!";
foreach my $dir (sort keys %dirs) {
    my $full_dir = $dir eq '' ? '..\\..\\..\\' : "..\\..\\..\\$dir";
    print $dir_fh "$full_dir\n";
}
close($dir_fh);

# Write file paths
open(my $file_fh, '>', $file_output) 
  or die "Cannot open file list output '$file_output': $!";
print $file_fh "$_\n" for sort @files;
close($file_fh);

print "Operation completed.\n";
print "- Directory paths written to $dir_output\n";
print "- File list written to $file_output\n";