class ContactData:

    def __init__(
        self,
        deliver_via_paper: bool,
        email: str,
        pli_id: int,
        first_name: str,
        last_name: str,
    ):
        self.deliver_via_paper: bool = deliver_via_paper
        self.email: str = email
        self.pli_id: int = pli_id
        self.first_name: str = first_name
        self.last_name: str = last_name
