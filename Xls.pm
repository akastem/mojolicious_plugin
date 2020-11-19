package App::Plugin::Xls;

use Mojo::Base 'Mojolicious::Plugin';

use Spreadsheet::WriteExcel::Simple;

sub register {
	my ($self, $app) = @_;
	
	$app->types->type(xls => 'application/vnd.ms-excel');
	
	$app->renderer->add_handler('xls' => sub {
		my ($renderer, $c, $output, $options) = @_;
		
		my $object   = $c->stash->{'xls'};
		my $filename = $c->stash->{'filename'};
		
		delete $options->{'encoding'};
		$options->{'format'} = 'xls';
		
		if ( $filename ) {
			my $headers = Mojo::Headers->new();
			   $headers->add( 'Content-Type',        'application/x-download;name=' . $filename );
			   $headers->add( 'Content-Disposition', 'attachment;filename=' . $filename );
			$c->res->content->headers($headers);
		}
		
		if (ref $object) {
			$$output = $object->data;
			return 1;
		}
		
		my $ss       = Spreadsheet::WriteExcel::Simple->new;
		my $heading  = $c->stash->{'heading'};
		my $result   = $c->stash->{'result'};
		my $settings = $c->stash->{'settings'};
		
		$ss->write_bold_row($heading) if (ref $heading);
		$ss->write_row($_) foreach (@$result);
		
		if (ref $settings) {
			$c->render_exception("invalid column width") unless defined $settings->{column_width};
			
			for my $col (keys %{$settings->{column_width}}) {
				$ss->sheet->set_column($col, $settings->{column_width}->{$col});
			}
		}
		
		$$output = $ss->data;
		
		return 1;
	});
}

1;
