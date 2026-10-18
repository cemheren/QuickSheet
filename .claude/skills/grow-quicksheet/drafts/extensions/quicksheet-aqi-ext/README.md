# quicksheet-aqi-ext

Live air quality on your QuickSheet desktop — US AQI, PM2.5 and PM10 for any city, powered by [Open-Meteo](https://open-meteo.com/) (no API key required).

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-aqi-ext
```

## Usage

```
aqi: London
aqi: New York
aqi: Delhi
aqi: 37.77,-122.42
```

Pass a city name, or `lat,lon` coordinates directly.

## What you see

```
┌─────────────────────────────────────┐
│ ● Delhi, IN                         │
│ US AQI 168 (Unhealthy)              │
│ PM2.5 88.4 ug/m3                    │
│ PM10 142.0 ug/m3                    │
└─────────────────────────────────────┘
```

- City name + country code
- US EPA AQI value and category (Good → Hazardous)
- PM2.5 and PM10 particulate concentrations (µg/m³)

Pin it next to [quicksheet-weather-ext](https://github.com/cemheren/quicksheet-weather-ext) for an ambient air + weather row on your wallpaper — handy during wildfire season or in high-pollution cities.

## How it works

Two keyless, free [Open-Meteo](https://open-meteo.com/) endpoints:

1. **Geocoding API** resolves the city name to latitude/longitude.
2. **Air Quality API** returns the current US AQI, PM2.5 and PM10 for those coordinates.

No API key, no sign-up. Open-Meteo is free for non-commercial use. Coordinates can be passed directly to skip geocoding.

### US AQI categories

| AQI | Category |
|-----|----------|
| 0–50 | Good |
| 51–100 | Moderate |
| 101–150 | Unhealthy (sensitive groups) |
| 151–200 | Unhealthy |
| 201–300 | Very Unhealthy |
| 301+ | Hazardous |

## Protocol

Implements the QuickSheet extension protocol (JSON-lines over stdin/stdout): responds to `init` with a `register` message for the `aqi` prefix, and to `activate` with a `write` of cell rows. Zero NuGet dependencies — .NET 9 stdlib only.

## License

MIT — see [LICENSE](LICENSE).
