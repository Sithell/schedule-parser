from dataclasses import dataclass


@dataclass
class Event:
    id: str
    status: str
    html_link: str
    summary: str
    description: str
    location: str
    color_id: str
    start: dict
    end: dict

    def asdict(self):
        return {
            'id': self.id,
            'status': self.status,
            'html_link': self.html_link,
            'summary': self.summary,
            'description': self.description,
            'location': self.location,
            'color_id': self.color_id,
            'start': self.start,
            'end': self.end,
        }
