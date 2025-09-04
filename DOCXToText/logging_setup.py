import logging


def configure_logging(verbosity: int) -> None:
	"""
	Configure root logging based on verbosity level.
	- 0: WARNING
	- 1: INFO
	- 2+: DEBUG
	"""
	level = logging.WARNING
	if verbosity == 1:
		level = logging.INFO
	elif verbosity >= 2:
		level = logging.DEBUG

	logging.basicConfig(
		level=level,
		format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
		datefmt="%Y-%m-%d %H:%M:%S",
	)


